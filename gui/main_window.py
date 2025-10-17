import logging
import os
import re
import threading
from typing import Any, Dict, Iterable, List, Optional, Set, Tuple

from PySide6.QtWidgets import (
    QMainWindow,
    QWidget,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QSplitter,
    QComboBox,
)
from PySide6.QtCore import Qt, QTimer, QUrl, Signal
from PySide6.QtGui import QAction, QActionGroup, QDesktopServices

from logic.project_manager import ProjectManager
from updater import (
    APP_VERSION,
    AUTHOR,
    RELEASE_DATE,
    check_for_updates,
    check_for_updates_background,
)
from gui.language_pair import LanguagePairWidget
from gui.drop_areas import DropArea
from gui.panels.left_panel import create_left_panel
from gui.panels.right_panel import create_right_panel
from gui.project_manager_dialog import ProjectManagerDialog
from gui.project_setup_widget import ProjectSetupWidget
from gui.styles import APP_STYLE
from gui.utils import format_language_display
from gui.rates_manager_window import RatesManagerWindow
from logic import rates_importer
from logic.user_config import add_language, load_languages
from logic.importers import import_project_info, import_xml_reports
from logic.service_config import ServiceConfig
from logic.pm_store import load_pm_history, save_pm_history
from logic.language_pairs import LanguagePairsMixin
from logic.legal_entities import get_legal_entity_metadata, load_legal_entities
from logic.translation_config import tr
from logic.logging_utils import get_last_run_log_path
from logic.xml_parser_common import language_identity, resolve_language_display
from logic.calculations import (
    convert_to_rub as calculate_convert_to_rub,
    on_currency_changed as calculate_on_currency_changed,
    set_currency_code as calculate_set_currency_code,
    update_total as calculate_update_total,
)

logger = logging.getLogger(__name__)


class TranslationCostCalculator(QMainWindow, LanguagePairsMixin):
    update_available = Signal(str)

    def __init__(self):
        super().__init__()
        self.language_pairs: Dict[str, LanguagePairWidget] = {}
        self.pair_headers: Dict[str, str] = {}
        self._pair_language_inputs: Dict[str, Dict[str, Dict[str, Any]]] = {}
        self._language_variant_regions: Dict[Tuple[str, str], Set[str]] = {}
        self.lang_display_ru: bool = True  # Controls language for quotation/Excel
        self.gui_lang: str = "ru"  # Controls application GUI language
        self._languages: List[Dict[str, str]] = load_languages()
        self.pm_managers, self.pm_last_index = load_pm_history()
        if 0 <= self.pm_last_index < len(self.pm_managers):
            self.current_pm = self.pm_managers[self.pm_last_index]
        else:
            self.current_pm = {"name_ru": "", "name_en": "", "email": ""}
        self.only_new_repeats_mode = False
        self.legal_entities = load_legal_entities()
        self.legal_entity_meta = get_legal_entity_metadata()
        self.currency_symbol = ""
        self.rates_window: Optional[RatesManagerWindow] = None
        self._import_pair_map: Dict[Tuple[str, str], str] = {}
        # Create labels early so slots triggered during initialization
        # (e.g. vat spin value changes) can safely update them.
        self.total_label = QLabel()
        self.discount_total_label = QLabel()
        self.markup_total_label = QLabel()
        self.project_manager = ProjectManager(self)
        self.setup_ui()
        self.setup_style()
        self.update_available.connect(self._handle_background_update_available)
        QTimer.singleShot(0, self.auto_check_for_updates)

    def setup_ui(self):
        self.setGeometry(100, 100, 1000, 600)
        self.setMinimumSize(600, 400)
        self.resize(1000, 650)
        self.update_title()

        self._build_menus()

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QHBoxLayout()
        splitter = QSplitter(Qt.Horizontal)

        left_panel = create_left_panel(self)
        right_panel = create_right_panel(self)

        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)

        splitter.setSizes([600, 960])

        main_layout.addWidget(splitter)
        central_widget.setLayout(main_layout)

    # Menu construction -------------------------------------------------
    def _build_menus(self) -> None:
        lang = self.gui_lang
        self.project_menu = self._create_project_menu(lang)
        self.export_menu = self._create_export_menu(lang)
        self.rates_menu = self._create_rates_menu(lang)
        self.pm_action = self._make_action(tr("Проджект менеджер", lang), self.show_pm_dialog)
        self.menuBar().addAction(self.pm_action)
        self.update_menu = self._create_update_menu(lang)
        self.about_action = self._make_action(tr("О программе", lang), self.show_about_dialog)
        self.menuBar().addAction(self.about_action)
        self._configure_language_menu(lang)
        self.update_menu_texts()

    def _create_project_menu(self, lang: str):
        menu = self.menuBar().addMenu(tr("Проект", lang))
        self.save_action = self._make_action(tr("Сохранить проект", lang), self.project_manager.save_project)
        self.load_action = self._make_action(tr("Загрузить проект", lang), self.project_manager.load_project)
        self.clear_action = self._make_action(tr("Очистить", lang), self.clear_all_data)
        self.open_log_action = self._make_action(tr("Открыть лог", lang), self.open_last_run_log)
        for action in (self.save_action, self.load_action, self.clear_action, self.open_log_action):
            menu.addAction(action)
        return menu

    def _create_export_menu(self, lang: str):
        menu = self.menuBar().addMenu(tr("Экспорт", lang))
        self.save_excel_action = self._make_action(tr("Сохранить Excel", lang), self.project_manager.save_excel)
        self.save_pdf_action = self._make_action(tr("Сохранить PDF", lang), self.project_manager.save_pdf)
        for action in (self.save_excel_action, self.save_pdf_action):
            menu.addAction(action)
        return menu

    def _create_rates_menu(self, lang: str):
        menu = self.menuBar().addMenu(tr("Ставки", lang))
        self.import_rates_action = self._make_action(tr("Открыть панель ставок", lang), self.open_rates_panel)
        menu.addAction(self.import_rates_action)
        return menu

    def _create_update_menu(self, lang: str):
        menu = self.menuBar().addMenu(tr("Обновление", lang))
        self.check_updates_action = self._make_action(tr("Проверить обновления", lang), self.manual_update_check)
        menu.addAction(self.check_updates_action)
        return menu

    def _configure_language_menu(self, lang: str) -> None:
        self.language_menu = self.menuBar().addMenu("Lang")
        self.language_menu.setToolTip(tr("Язык", lang))
        self.language_menu.menuAction().setToolTip(tr("Язык", lang))
        self.lang_action_group = QActionGroup(self)
        self.lang_ru_action = QAction("русский", self)
        self.lang_en_action = QAction("english", self)
        for action, code in ((self.lang_ru_action, "ru"), (self.lang_en_action, "en")):
            action.setCheckable(True)
            action.triggered.connect(lambda checked, value=code: self.set_app_language(value))
            self.lang_action_group.addAction(action)
            self.language_menu.addAction(action)
        self.lang_ru_action.setChecked(self.gui_lang == "ru")
        self.lang_en_action.setChecked(self.gui_lang != "ru")

    def _make_action(self, text: str, slot) -> QAction:
        action = QAction(text, self)
        action.triggered.connect(slot)
        return action

    def open_last_run_log(self):
        """Open the last run log file in the system text editor."""
        log_path = get_last_run_log_path()
        try:
            log_path.touch(exist_ok=True)
        except OSError:
            logger.exception("Failed to touch log file: %s", log_path)
            QMessageBox.warning(
                self,
                tr("Ошибка", self.gui_lang),
                tr("Не удалось получить доступ к файлу журнала.", self.gui_lang),
            )
            return

        opened = QDesktopServices.openUrl(QUrl.fromLocalFile(str(log_path)))
        if not opened:
            logger.warning("System was unable to open log file: %s", log_path)
            QMessageBox.warning(
                self,
                tr("Ошибка", self.gui_lang),
                tr("Не удалось открыть файл журнала.", self.gui_lang),
            )

    def _make_lang_combo(self) -> QComboBox:
        cb = QComboBox()
        self.populate_lang_combo(cb)
        return cb

    def populate_lang_combo(self, combo: QComboBox):
        prev_text = combo.currentText() if combo.isEditable() else ""
        prev_idx = combo.currentIndex()
        prev_obj = combo.itemData(prev_idx) if prev_idx >= 0 else None
        typed_over_selection = False
        if combo.isEditable() and prev_idx >= 0 and prev_text:
            current_item_text = combo.itemText(prev_idx)
            if current_item_text.strip() != prev_text.strip():
                typed_over_selection = True

        combo.blockSignals(True)
        combo.clear()
        for lang in self._languages:
            name = lang["ru"] if self.lang_display_ru else lang["en"]
            locale = "ru" if self.lang_display_ru else "en"
            name = self._prepare_language_label(name, locale)
            label = f"{name}"
            combo.addItem(label, lang)
        combo.blockSignals(False)

        if isinstance(prev_obj, dict) and not typed_over_selection:
            for i in range(combo.count()):
                d = combo.itemData(i)
                if (
                    isinstance(d, dict)
                    and d.get("en") == prev_obj.get("en")
                    and d.get("ru") == prev_obj.get("ru")
                ):
                    combo.setCurrentIndex(i)
                    break
        elif prev_text:
            combo.setEditable(True)
            combo.setCurrentIndex(-1)
            combo.setEditText(prev_text)

    def set_app_language(self, lang: str):
        """Change application GUI language via menu action."""
        self.gui_lang = lang
        self._update_gui_language(lang)
        self.lang_ru_action.setChecked(lang == "ru")
        self.lang_en_action.setChecked(lang == "en")
        desired_slider_value = 1 if lang == "ru" else 0
        if self.lang_mode_slider.value() != desired_slider_value:
            self.lang_mode_slider.setValue(desired_slider_value)

    def on_lang_mode_changed(self, value: int):
        """Handle slider changes – update language pair names everywhere."""
        lang = "ru" if value == 1 else "en"
        self.lang_display_ru = value == 1
        self._update_language_names(lang)
        self.update_total()

    def _update_language_names(self, lang: str):
        """Update language names in GUI widgets and Excel headers."""
        self.populate_lang_combo(self.source_lang_combo)
        self.populate_lang_combo(self.target_lang_combo)
        if getattr(self, "project_setup_widget", None):
            self.project_setup_widget.set_language(lang)
        if getattr(self, "additional_services_widget", None):
            self.additional_services_widget.set_language(lang)
        for pair_key, widget in self.language_pairs.items():
            widget.set_language(lang)
            display_name = self._display_pair_name(pair_key)
            widget.set_pair_name(display_name)
            _, right_key = self._extract_pair_parts(pair_key)
            target_entry = self._pair_language_inputs.get(pair_key, {}).get("target")
            if target_entry:
                labels = self._labels_from_entry(target_entry)
                self.pair_headers[pair_key] = labels[lang]
            elif right_key:
                lang_info = self._find_language_by_key(right_key)
                self.pair_headers[pair_key] = lang_info[lang]
            else:
                self.pair_headers[pair_key] = display_name
        self.update_pairs_list()

    def _update_gui_language(self, lang: str):
        """Update visible GUI texts when language is changed via menu."""
        self.project_group.setTitle(tr("Информация о проекте", lang))
        self.project_name_label.setText(tr("Название проекта", lang) + ":")
        self.client_name_label.setText(tr("Название клиента", lang) + ":")
        self.contact_person_label.setText(tr("Контактное лицо", lang) + ":")
        self.email_label.setText(tr("Email", lang) + ":")
        self.legal_entity_label.setText(tr("Юрлицо", lang) + ":")
        self.legal_entity_placeholder = tr("Выберите юрлицо", lang)
        if self.legal_entity_combo.count() > 0:
            self.legal_entity_combo.setItemText(0, self.legal_entity_placeholder)
        self.currency_label.setText(tr("Валюта", lang) + ":")
        self.currency_placeholder = tr("Выберите валюту", lang)
        if self.currency_combo.count() > 0:
            self.currency_combo.setItemText(0, self.currency_placeholder)
        self.convert_btn.setText(tr("Конвертировать в рубли", lang))
        self.vat_label.setText(tr("НДС, %", lang) + ":")
        self.language_names_label.setText(tr("Названия языков", lang) + ":")
        self.add_pair_btn.setText(tr("Добавить языковую пару", lang))
        self.current_pairs_label.setText(tr("Текущие пары", lang) + ":")
        self.clear_pairs_btn.setText(tr("Очистить", lang))
        self.project_setup_label.setText(tr("Запуск и управление проектом", lang) + ":")
        self.add_lang_group.setTitle(tr("Добавить язык в справочник", lang))
        self.lang_ru_label.setText(tr("Название RU", lang) + ":")
        self.lang_en_label.setText(tr("Название EN", lang) + ":")
        self.btn_add_lang.setText(tr("Добавить язык", lang))
        self.new_lang_ru.setPlaceholderText(tr("Литовский", "ru"))
        self.new_lang_en.setPlaceholderText(tr("Lithuanian", "en"))

        self.pairs_group.setTitle(tr("Языковые пары", lang))
        self.tabs.setTabText(0, tr("Языковые пары", lang))
        self.tabs.setTabText(1, tr("Дополнительные услуги", lang))
        if getattr(self, "delete_all_pairs_btn", None):
            self.delete_all_pairs_btn.setText(tr("Удалить все языки", lang))
        if getattr(self, "drop_hint_label", None):
            self.drop_hint_label.setText(
                tr(
                    "Перетащите XML файлы отчетов Trados или Smartcat сюда для автоматического заполнения",
                    lang,
                )
            )
        if self.only_new_repeats_mode:
            self.only_new_repeats_btn.setText(tr("Показать 4 строки", lang))
        else:
            self.only_new_repeats_btn.setText(
                tr("Только новые слова и повторы", lang)
            )
        self.update_title()
        if self.rates_window:
            self.rates_window.set_language(lang)
        self.update_menu_texts()

    def update_menu_texts(self):
        lang = self.gui_lang
        self.project_menu.setTitle(tr("Проект", lang))
        self.save_action.setText(tr("Сохранить проект", lang))
        self.load_action.setText(tr("Загрузить проект", lang))
        self.clear_action.setText(tr("Очистить", lang))
        self.open_log_action.setText(tr("Открыть лог", lang))
        self.export_menu.setTitle(tr("Экспорт", lang))
        self.save_excel_action.setText(tr("Сохранить Excel", lang))
        self.save_pdf_action.setText(tr("Сохранить PDF", lang))
        self.rates_menu.setTitle(tr("Импорт ставок", lang))
        self.import_rates_action.setText(tr("Импортировать из Excel", lang))
        self.pm_action.setText(tr("Проджект менеджер", lang))
        self.update_menu.setTitle(tr("Обновление", lang))
        self.check_updates_action.setText(tr("Проверить обновления", lang))
        self.about_action.setText(tr("О программе", lang))
        self.language_menu.setTitle("Lang")
        self.language_menu.setToolTip(tr("Язык", lang))
        self.language_menu.menuAction().setText("Lang")
        self.language_menu.menuAction().setToolTip(tr("Язык", lang))
        self.lang_ru_action.setText("русский")
        self.lang_en_action.setText("english")

    def manual_update_check(self):
        check_for_updates(self, force=True)

    def auto_check_for_updates(self):
        def worker():
            try:
                version = check_for_updates_background(force=False)
            except Exception:  # pragma: no cover - network errors, etc.
                logger.debug("Background update check failed", exc_info=True)
                return
            if version:
                self.update_available.emit(version)

        threading.Thread(target=worker, daemon=True).start()

    def _handle_background_update_available(self, version: str):
        lang = self.gui_lang
        reply = QMessageBox.question(
            self,
            tr("Доступно обновление", lang),
            tr("Доступна новая версия {0}. Проверить обновление сейчас?", lang).format(version),
        )
        if reply == QMessageBox.Yes:
            self.manual_update_check()

    def show_about_dialog(self):
        lang = self.gui_lang
        text = (
            f"{tr('Версия', lang)}: {APP_VERSION}\n"
            f"{tr('Дата', lang)}: {RELEASE_DATE}\n"
            f"{tr('Автор', lang)}: {AUTHOR}"
        )
        QMessageBox.information(self, tr("О программе", lang), text)

    def on_currency_index_changed(self, index: int):
        code = self.currency_combo.itemText(index) if index > 0 else ""
        self.on_currency_changed(code)

    def get_current_currency_code(self) -> str:
        index = self.currency_combo.currentIndex()
        if index <= 0:
            return ""
        return self.currency_combo.itemText(index)

    def set_currency_code(self, code: Optional[str]) -> bool:
        return calculate_set_currency_code(self, code)

    def on_currency_changed(self, code: str):
        calculate_on_currency_changed(self, code)

    def get_selected_legal_entity(self) -> str:
        """Return currently selected legal entity or empty string if none."""
        idx = self.legal_entity_combo.currentIndex()
        if idx <= 0:
            return ""
        return self.legal_entity_combo.currentText()

    def convert_to_rub(self):
        calculate_convert_to_rub(self)

    def on_legal_entity_changed(self, entity: str):
        if entity == self.legal_entity_placeholder:
            entity = ""

        normalized_entity = entity.strip() if entity else ""
        if (
            normalized_entity.lower() == "logrus it"
            and self.lang_mode_slider.value() == 1
        ):
            self.lang_mode_slider.setValue(0)

        meta = self.legal_entity_meta.get(entity, {}) if entity else {}
        vat_enabled = bool(meta.get("vat_enabled"))
        self.vat_spin.setEnabled(vat_enabled)
        if vat_enabled:
            default_vat = meta.get("default_vat", 0.0)
            try:
                default_vat_value = float(default_vat)
            except (TypeError, ValueError):
                default_vat_value = 0.0
            if self.vat_spin.value() == 0.0 and default_vat_value > 0:
                self.vat_spin.setValue(default_vat_value)
        else:
            self.vat_spin.setValue(0.0)

    def setup_drag_drop(self):
        drop_area = DropArea(self.handle_xml_drop, lambda: self.gui_lang)

        drop_area.setWidget(self.pairs_container_widget)

        self.tabs.removeTab(0)
        lang = self.gui_lang
        self.tabs.insertTab(0, drop_area, tr("Языковые пары", lang))
        self.pairs_scroll = drop_area

    def handle_project_info_drop(self, paths: List[str]):
        lang = self.gui_lang
        result, errors = import_project_info(paths)

        if errors:
            QMessageBox.warning(
                self,
                tr("Ошибка обработки Outlook файла", lang),
                "\n".join(errors),
            )

        if not result:
            return

        self._apply_project_info_payload(result)

    def _apply_project_info_payload(self, payload: Dict[str, Any]):
        lang = self.gui_lang
        data = payload.get("data", {})

        updated_fields: List[str] = []
        manual_checks: List[str] = list(payload.get("warnings", []))

        def update_field(widget, value: Optional[str], label_key: str):
            if value:
                widget.setText(value)
                updated_fields.append(tr(label_key, lang))
            else:
                widget.clear()

        update_field(self.project_name_edit, data.get("project_name"), "Название проекта")
        update_field(self.client_name_edit, data.get("client_name"), "Название клиента")
        update_field(self.contact_person_edit, data.get("contact_name"), "Контактное лицо")
        update_field(self.email_edit, data.get("email"), "Email")

        legal_entity_value = (data.get("legal_entity") or "").strip()
        if payload.get("flags", {}).get("force_ru_mode") and self.lang_mode_slider.value() != 0:
            self.lang_mode_slider.setValue(0)

        if legal_entity_value:
            idx = self.legal_entity_combo.findText(
                legal_entity_value, Qt.MatchFixedString
            )
            if idx < 0:
                target = legal_entity_value.lower()
                for i in range(self.legal_entity_combo.count()):
                    text = self.legal_entity_combo.itemText(i).strip().lower()
                    if text == target:
                        idx = i
                        break
            if idx >= 0:
                self.legal_entity_combo.setCurrentIndex(idx)
                updated_fields.append(tr("Юрлицо", lang))
            else:
                self.legal_entity_combo.setCurrentIndex(0)
                manual_checks.append(f"{tr('Юрлицо', lang)}: {legal_entity_value}")
        else:
            self.legal_entity_combo.setCurrentIndex(0)

        currency_code = data.get("currency_code")
        if currency_code:
            if self.set_currency_code(currency_code):
                updated_fields.append(tr("Валюта", lang))
            else:
                self.set_currency_code(None)
                manual_checks.append(f"{tr('Валюта', lang)}: {currency_code}")
        else:
            self.set_currency_code(None)

        missing = [tr(name, lang) for name in payload.get("missing_fields", [])]

        sender = payload.get("sender", {})
        sender_parts: List[str] = []
        if sender.get("name"):
            sender_parts.append(sender["name"])
        if sender.get("email"):
            sender_parts.append(sender["email"])

        def format_section(title: str, values: List[str]) -> Optional[str]:
            unique_values = list(dict.fromkeys(v for v in values if v))
            if not unique_values:
                return None
            bullets = "\n  • ".join(unique_values)
            return f"{title}:\n  • {bullets}"

        subject = (payload.get("subject") or "").strip()
        source_name = os.path.basename(payload.get("source_path", ""))
        if subject:
            header = f"{tr('Outlook письмо', lang)}: {subject}"
            if source_name and source_name not in subject:
                header += f"\n{tr('Файл', lang)}: {source_name}"
        else:
            header = f"{tr('Outlook письмо', lang)}: {source_name}"

        message_sections: List[str] = [header]

        section = format_section(tr("Обновлены поля", lang), updated_fields)
        if section:
            message_sections.append(section)

        section = format_section(tr("Не удалось определить", lang), missing)
        if section:
            message_sections.append(section)

        section = format_section(tr("Проверьте вручную", lang), manual_checks)
        if section:
            message_sections.append(section)

        section = format_section(tr("Отправитель", lang), sender_parts)
        if section:
            message_sections.append(section)

        sent_at = sender.get("sent_at")
        if sent_at:
            message_sections.append(f"{tr('Дата отправки', lang)}: {sent_at}")

        QMessageBox.information(
            self, tr("Готово", lang), "\n\n".join(message_sections)
        )

    def _hide_drop_hint(self):
        if getattr(self, "drop_hint_label", None):
            self.pairs_layout.removeWidget(self.drop_hint_label)
            self.drop_hint_label.deleteLater()
            self.drop_hint_label = None
        if isinstance(self.pairs_scroll, DropArea):
            self.pairs_scroll.disable_hint_style()

    def toggle_only_new_repeats_mode(self):
        self.only_new_repeats_mode = not self.only_new_repeats_mode
        for w in self.language_pairs.values():
            w.set_only_new_and_repeats_mode(self.only_new_repeats_mode)
        lang = self.gui_lang
        if self.only_new_repeats_mode:
            self.only_new_repeats_btn.setText(tr("Показать 4 строки", lang))
        else:
            self.only_new_repeats_btn.setText(
                tr("Только новые слова и повторы", lang)
            )

    def setup_style(self):
        self.setStyleSheet(APP_STYLE)

    def update_title(self):
        lang = getattr(self, "gui_lang", "ru")
        name = ""
        preferred_keys = ("name_en", "name_ru") if lang == "en" else ("name_ru", "name_en")
        for key in preferred_keys:
            name = (self.current_pm or {}).get(key) or ""
            if name:
                break
        if name:
            self.setWindowTitle(f"RateApp - {name}")
        else:
            self.setWindowTitle("RateApp")

    def show_pm_dialog(self):
        lang = self.gui_lang
        dlg = ProjectManagerDialog(self.pm_managers, self.pm_last_index, lang, self)
        if dlg.exec():
            self.pm_managers, self.pm_last_index = dlg.result()
            save_pm_history(self.pm_managers, self.pm_last_index)
            if 0 <= self.pm_last_index < len(self.pm_managers):
                self.current_pm = self.pm_managers[self.pm_last_index]
            else:
                self.current_pm = {"name_ru": "", "name_en": "", "email": ""}
            self.update_title()

    def update_project_setup_volume_from_spin(self, value: float):
        if getattr(self, "project_setup_widget", None):
            self.project_setup_widget.set_volume(value)

    def on_project_setup_item_changed(self, item):
        if item.row() == 0 and item.column() == 1:
            try:
                val = float(item.text() or "0")
            except ValueError:
                val = 0
            self.project_setup_fee_spin.blockSignals(True)
            self.project_setup_fee_spin.setValue(val)
            self.project_setup_fee_spin.blockSignals(False)

    def remove_project_setup_widget(self):
        if getattr(self, "project_setup_widget", None):
            self.project_setup_widget.setParent(None)
            self.project_setup_widget = None
        self.update_total()

    def update_total(self, *_):
        calculate_update_total(self)

    def handle_add_language(self):
        ru = (self.new_lang_ru.text() or "").strip()
        en = (self.new_lang_en.text() or "").strip()

        if not (ru or en):
            lang = self.gui_lang
            QMessageBox.warning(
                self,
                tr("Ошибка", lang),
                tr("Укажите хотя бы одно название (RU или EN).", lang),
            )
            return

        if add_language(en, ru):
            lang = self.gui_lang
            QMessageBox.information(
                self, tr("Готово", lang), tr("Язык сохранён в конфиг.", lang)
            )
            self._languages = load_languages()
            self.populate_lang_combo(self.source_lang_combo)
            self.populate_lang_combo(self.target_lang_combo)
            self.new_lang_ru.clear()
            self.new_lang_en.clear()
        else:
            lang = self.gui_lang
            QMessageBox.critical(
                self, tr("Ошибка", lang), tr("Не удалось сохранить язык в конфиг.", lang)
            )

    def _parse_combo(self, combo: QComboBox) -> Dict[str, Any]:
        text = combo.currentText().strip()
        idx = combo.currentIndex()
        if idx >= 0 and text == combo.itemText(idx):
            data = combo.itemData(idx)
            if isinstance(data, dict):
                return {
                    "en": data.get("en", ""),
                    "ru": data.get("ru", ""),
                    "text": text,
                    "dict": True,
                }
        return {"en": "", "ru": "", "text": text, "dict": False}

    def _store_pair_language_inputs(
        self,
        pair_key: str,
        source: Dict[str, Any],
        target: Dict[str, Any],
        left_key: Optional[str] = None,
        right_key: Optional[str] = None,
    ) -> None:
        source_entry = dict(source or {})
        target_entry = dict(target or {})
        if left_key and not source_entry.get("key"):
            source_entry["key"] = left_key
        if right_key and not target_entry.get("key"):
            target_entry["key"] = right_key
        if "key" not in source_entry:
            source_entry["key"] = (
                source_entry.get("text")
                or source_entry.get("en")
                or source_entry.get("ru")
                or ""
            )
        if "key" not in target_entry:
            target_entry["key"] = (
                target_entry.get("text")
                or target_entry.get("en")
                or target_entry.get("ru")
                or ""
            )
        self._pair_language_inputs[pair_key] = {
            "source": source_entry,
            "target": target_entry,
        }

    def _pair_sort_key(self, pair_key: str) -> str:
        _, right = self._extract_pair_parts(pair_key)
        return right or pair_key

    def _update_language_variant_regions_from_pairs(
        self, pair_keys: Iterable[str]
    ) -> None:
        region_map: Dict[Tuple[str, str], Set[str]] = {}
        for pair_key in pair_keys:
            left, right = self._extract_pair_parts(pair_key)
            for value in (left, right):
                language_code, script_code, territory_code = language_identity(value)
                if not language_code or not territory_code:
                    continue
                map_key = (language_code, script_code)
                region_map.setdefault(map_key, set()).add(territory_code)
        self._language_variant_regions = region_map

    def _prepare_language_label(self, name: str, locale: str) -> str:
        if not name:
            return ""

        formatted = format_language_display(name, locale)
        language_code, script_code, territory_code = language_identity(name)
        if territory_code:
            territories = self._language_variant_regions.get((language_code, script_code))
            if not territories or len(territories) <= 1:
                formatted = re.sub(r"\s*\([^()]*\)\s*$", "", formatted).strip()
        return formatted

    def _find_language_by_key(self, key: str) -> Dict[str, str]:
        norm = key.strip().lower()
        for lang in self._languages:
            if norm == lang["en"].lower() or norm == lang["ru"].lower():
                return {
                    "en": self._prepare_language_label(lang["en"], "en"),
                    "ru": self._prepare_language_label(lang["ru"], "ru"),
                }

        ru_name = resolve_language_display(key, locale="ru")
        en_name = resolve_language_display(key, locale="en")

        if not en_name:
            en_name = key
        if not ru_name:
            ru_name = key

        en_name = self._prepare_language_label(en_name, "en")
        ru_name = self._prepare_language_label(ru_name, "ru")
        return {"en": en_name, "ru": ru_name}

    def update_pairs_list(self):
        self.pairs_list.setText(
            "\n".join(
                w.pair_name
                for key, w in sorted(
                    self.language_pairs.items(),
                    key=lambda kv: self._pair_sort_key(kv[0]),
                )
            )
        )
        pair_count = len(self.language_pairs)
        lang = self.gui_lang
        if getattr(self, "delete_all_pairs_btn", None):
            self.delete_all_pairs_btn.setEnabled(pair_count > 0)
        self.language_pairs_count_label.setText(
            f"{tr('Загружено языковых пар', lang)}: {pair_count}"
        )
        auto_fee = max(0.5, round((pair_count + 1) / 4 * 4) / 4)
        self.project_setup_fee_spin.blockSignals(True)
        self.project_setup_fee_spin.setValue(auto_fee)
        self.project_setup_fee_spin.blockSignals(False)
        self.update_project_setup_volume_from_spin(auto_fee)

    def _on_widget_remove_requested(self, widget: LanguagePairWidget):
        for key, w in list(self.language_pairs.items()):
            if w is widget:
                self.remove_language_pair(key)
                break

    def delete_all_language_pairs(self) -> None:
        if not self.language_pairs:
            return
        lang = self.gui_lang
        reply = QMessageBox.question(
            self,
            tr("Подтверждение", lang),
            tr("Вы уверены, что хотите удалить все языки?", lang),
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return
        self.clear_language_pairs()

    def on_pair_name_changed(self, widget: LanguagePairWidget, new_name: str):
        for key, w in list(self.language_pairs.items()):
            if w is widget:
                if (
                    new_name in self.language_pairs
                    and self.language_pairs[new_name] is not widget
                ):
                    # revert to old name if duplicate
                    widget.set_pair_name(key)
                    return
                header_title = self.pair_headers.pop(key, widget.pair_name)
                self.language_pairs.pop(key)
                self.language_pairs[new_name] = widget
                self.pair_headers[new_name] = header_title
                lang_inputs = self._pair_language_inputs.pop(key, None)
                if lang_inputs is not None:
                    self._pair_language_inputs[new_name] = lang_inputs
                break
        self.update_pairs_list()

    def clear_all_data(self):
        """Reset all user-entered and loaded data."""
        self.project_name_edit.clear()
        self.client_name_edit.clear()
        self.contact_person_edit.clear()
        self.email_edit.clear()
        self.legal_entity_combo.setCurrentIndex(0)
        self.currency_combo.setCurrentIndex(0)
        self.vat_spin.setValue(20.0)

        if getattr(self, "lang_mode_slider", None) is not None:
            self.lang_mode_slider.setValue(1)
        if getattr(self, "source_lang_combo", None) is not None:
            self.source_lang_combo.setCurrentIndex(0)
            self.source_lang_combo.setEditText("")
        if getattr(self, "target_lang_combo", None) is not None:
            self.target_lang_combo.setCurrentIndex(0)
            self.target_lang_combo.setEditText("")
        if getattr(self, "new_lang_ru", None) is not None:
            self.new_lang_ru.clear()
        if getattr(self, "new_lang_en", None) is not None:
            self.new_lang_en.clear()

        est_lang = "ru" if self.lang_display_ru else "en"
        self.project_setup_fee_spin.setValue(0.5)
        if getattr(self, "project_setup_widget", None):
            self.project_setup_widget.load_data(
                [
                    {
                        "parameter": tr("Запуск и управление проектом", est_lang),
                        "volume": self.project_setup_fee_spin.value(),
                        "rate": 0.0,
                        "total": 0.0,
                    }
                ],
                enabled=True,
            )
            self.project_setup_widget.set_discount_percent(0.0)
            self.project_setup_widget.set_markup_percent(0.0)

        self.clear_language_pairs()
        if getattr(self, "additional_services_widget", None):
            self.additional_services_widget.load_data([])

        self._import_pair_map = {}
        if getattr(self, "rates_window", None):
            self.rates_window.reset_state()

        self.only_new_repeats_mode = False
        if getattr(self, "only_new_repeats_btn", None):
            self.only_new_repeats_btn.setText(
                tr("Только новые слова и повторы", self.gui_lang)
            )

        self.update_total()

    def open_rates_panel(self) -> None:
        if not self.language_pairs:
            lang = self.gui_lang
            QMessageBox.warning(
                self,
                tr("Ошибка", lang),
                tr("Сначала добавьте языковые пары", lang),
            )
            return

        pairs, pair_map = self._collect_pairs_for_rates()
        if not pairs:
            return

        self._import_pair_map = pair_map

        if self.rates_window is None:
            self.rates_window = RatesManagerWindow(self)
            self.rates_window.destroyed.connect(self._on_rates_window_destroyed)

        self.rates_window.update_pairs(pairs)
        self.rates_window.show()
        self.rates_window.raise_()
        self.rates_window.activateWindow()

    def import_rates_from_excel(self) -> None:
        """Backward-compatible alias for external integrations."""

        self.open_rates_panel()

    def _collect_pairs_for_rates(self) -> Tuple[List[Tuple[str, str]], Dict[Tuple[str, str], str]]:
        pairs: List[Tuple[str, str]] = []
        pair_map: Dict[Tuple[str, str], str] = {}
        lang_key = "ru" if self.lang_display_ru else "en"
        locale = "ru" if self.lang_display_ru else "en"
        for key in self.language_pairs:
            entries = self._pair_language_inputs.get(key, {})
            src_entry = entries.get("source") if isinstance(entries, dict) else None
            tgt_entry = entries.get("target") if isinstance(entries, dict) else None

            src_label = ""
            if isinstance(src_entry, dict):
                src_labels = self._labels_from_entry(src_entry)
                src_label = src_labels.get(lang_key, "")
            tgt_label = ""
            if isinstance(tgt_entry, dict):
                tgt_labels = self._labels_from_entry(tgt_entry)
                tgt_label = tgt_labels.get(lang_key, "")

            if not src_label or not tgt_label:
                parts = re.split(r"\s*(?:→|->|-|>)\s*", key, maxsplit=1)
                if len(parts) == 2:
                    fallback_src, fallback_tgt = parts
                else:
                    fallback_src, fallback_tgt = key, ""
                if not src_label:
                    src_label = self._prepare_language_label(fallback_src, locale)
                if not tgt_label and fallback_tgt:
                    tgt_label = self._prepare_language_label(fallback_tgt, locale)

            src_label = src_label.strip()
            tgt_label = tgt_label.strip()
            if not src_label or not tgt_label:
                continue

            pair = (src_label, tgt_label)
            pairs.append(pair)
            pair_map[pair] = key
        return pairs, pair_map

    def _on_rates_window_destroyed(self) -> None:
        self.rates_window = None
        self._import_pair_map = {}

    def _apply_rates_from_matches(
        self, matches: Iterable[rates_importer.PairMatch], rate_key: str
    ) -> None:
        for match in matches:
            if not match.rates:
                continue
            pair_key = self._import_pair_map.get((match.gui_source, match.gui_target))
            if not pair_key:
                continue
            widget = self.language_pairs.get(pair_key)
            if widget:
                value = match.rates.get(rate_key)
                if value is not None:
                    widget.set_basic_rate(value)

    def handle_xml_drop(self, paths: List[str], replace: bool = False):
        result, errors = import_xml_reports(paths)

        if errors:
            lang = self.gui_lang
            QMessageBox.critical(self, tr("Ошибка", lang), "\n".join(errors))
            return

        warnings = result.get("warnings", [])
        if warnings:
            lang = self.gui_lang
            warning_msg = (
                f"{tr('Предупреждения при обработке файлов', lang)}:\n"
                + "\n".join(warnings)
            )
            QMessageBox.warning(self, tr("Предупреждение", lang), warning_msg)

        data = result.get("data", {})
        report_sources = result.get("report_sources", {})

        combined_keys = set(self.language_pairs.keys()) | set(data.keys())
        self._update_language_variant_regions_from_pairs(combined_keys)

        if not data:
            lang = self.gui_lang
            message_lines = [
                tr("В XML файлах не найдено данных о языковых парах.", lang),
                "",
                tr("Возможные причины:", lang),
                tr("1. XML файлы имеют нестандартную структуру", lang),
                tr("2. Не найдены элементы LanguageDirection", lang),
                tr("3. Отсутствуют данные о языках или объемах", lang),
                "",
                tr("Проверьте консоль для детальной информации.", lang),
            ]
            QMessageBox.warning(
                self,
                tr("Результат обработки", lang),
                "\n".join(message_lines),
            )
            return

        self._hide_drop_hint()

        added_pairs = 0
        updated_pairs = 0

        for pair_key, volumes in sorted(
            data.items(), key=lambda kv: self._pair_sort_key(kv[0])
        ):
            widget = self.language_pairs.get(pair_key)
            sources_for_pair = report_sources.get(pair_key, [])

            left_raw, right_raw = self._extract_pair_parts(pair_key)
            self._store_pair_language_inputs(
                pair_key,
                {"text": left_raw, "dict": False, "key": left_raw},
                {"text": right_raw, "dict": False, "key": right_raw},
            )

            display_name = self._display_pair_name(pair_key)
            _, tgt_key = self._extract_pair_parts(pair_key)
            target_entry = self._pair_language_inputs.get(pair_key, {}).get("target")
            if target_entry:
                lang_labels = self._labels_from_entry(target_entry)
                header_title = (
                    lang_labels["ru"] if self.lang_display_ru else lang_labels["en"]
                )
            else:
                lang_info = self._find_language_by_key(tgt_key or pair_key)
                header_title = (
                    lang_info["ru"] if self.lang_display_ru else lang_info["en"]
                )

            if widget is None:
                widget = LanguagePairWidget(
                    display_name,
                    self.currency_symbol,
                    self.get_current_currency_code(),
                    lang="ru" if self.lang_display_ru else "en",
                )
                widget.remove_requested.connect(
                    lambda w=widget: self._on_widget_remove_requested(w)
                )
                widget.subtotal_changed.connect(self.update_total)
                widget.name_changed.connect(
                    lambda new_name, w=widget: self.on_pair_name_changed(w, new_name)
                )
                self.language_pairs[pair_key] = widget
                self.pairs_layout.insertWidget(
                    self.pairs_layout.count() - 1, widget
                )
                self.pair_headers[pair_key] = header_title
                added_pairs += 1
            else:
                widget.set_pair_name(display_name)
                widget.set_language("ru" if self.lang_display_ru else "en")
                self.pair_headers[pair_key] = header_title
                updated_pairs += 1

            widget.set_report_sources(sources_for_pair, replace=replace)

            if self.only_new_repeats_mode:
                widget.set_only_new_and_repeats_mode(True)

            group = widget.translation_group
            table = group.table
            group.setChecked(True)

            if self.only_new_repeats_mode:
                new_total = 0
                repeat_total = 0
                repeat_name = "Перевод, повторы и 100% совпадения (30%)"
                for row_name, add_val in volumes.items():
                    if row_name == repeat_name:
                        repeat_total += add_val
                    else:
                        new_total += add_val
                total_volume = new_total + repeat_total
                repeat_row = table.rowCount() - 1
                if replace:
                    table.item(0, 1).setText(str(new_total))
                    table.item(repeat_row, 1).setText(str(repeat_total))
                else:
                    try:
                        prev_new = float(
                            table.item(0, 1).text() if table.item(0, 1) else "0"
                        )
                    except ValueError:
                        prev_new = 0
                    try:
                        prev_rep = float(
                            table.item(repeat_row, 1).text()
                            if table.item(repeat_row, 1)
                            else "0"
                        )
                    except ValueError:
                        prev_rep = 0
                    table.item(0, 1).setText(str(prev_new + new_total))
                    table.item(repeat_row, 1).setText(str(prev_rep + repeat_total))
            else:
                total_volume = 0
                for idx, row_info in enumerate(ServiceConfig.TRANSLATION_ROWS):
                    row_name = row_info["name"]
                    add_val = volumes.get(row_name, 0)
                    total_volume += add_val

                    if replace:
                        table.item(idx, 1).setText(str(add_val))
                    else:
                        try:
                            prev_text = (
                                table.item(idx, 1).text()
                                if table.item(idx, 1)
                                else "0"
                            )
                            prev = float(prev_text or "0")
                            new_val = prev + add_val
                            table.item(idx, 1).setText(str(new_val))
                        except (ValueError, TypeError):
                            table.item(idx, 1).setText(str(add_val))

            widget.update_rates_and_sums(
                table, group.rows_config, group.base_rate_row
            )

        sorted_items = sorted(
            self.language_pairs.items(), key=lambda kv: self._pair_sort_key(kv[0])
        )
        for w in self.language_pairs.values():
            self.pairs_layout.removeWidget(w)
        for _, w in sorted_items:
            self.pairs_layout.insertWidget(self.pairs_layout.count() - 1, w)
        self.language_pairs = dict(sorted_items)

        self.update_pairs_list()
        self.update_total()

        self._update_language_variant_regions_from_pairs(self.language_pairs.keys())

        result_msg = "Обработка завершена!\n\n"
        if added_pairs > 0:
            result_msg += f"Добавлено новых языковых пар: {added_pairs}\n"
        if updated_pairs > 0:
            result_msg += f"Обновлено существующих пар: {updated_pairs}\n"
        result_msg += f"\nВсего обработано языковых пар: {len(data)}"


    def load_project_data(self, project_data: Dict[str, Any]):
        self.project_name_edit.setText(project_data.get("project_name", ""))
        self.client_name_edit.setText(project_data.get("client_name", ""))
        self.contact_person_edit.setText(project_data.get("contact_person", ""))
        self.email_edit.setText(project_data.get("email", ""))
        if hasattr(self, "legal_entity_combo"):
            le = project_data.get("legal_entity", "")
            if le:
                self.legal_entity_combo.setCurrentText(le)
            else:
                self.legal_entity_combo.setCurrentIndex(0)
        if hasattr(self, "currency_combo"):
            saved_currency = project_data.get("currency")
            self.set_currency_code(saved_currency)
        if hasattr(self, "vat_spin"):
            self.vat_spin.setValue(project_data.get("vat_rate", 20.0))
        self.on_legal_entity_changed(self.get_selected_legal_entity())

        self.only_new_repeats_mode = project_data.get("only_new_repeats_mode", False)
        if getattr(self, "only_new_repeats_btn", None):
            if self.only_new_repeats_mode:
                self.only_new_repeats_btn.setText(tr("Показать 4 строки", self.gui_lang))
            else:
                self.only_new_repeats_btn.setText(
                    tr("Только новые слова и повторы", self.gui_lang)
                )

        for w in self.language_pairs.values():
            w.setParent(None)
        self.language_pairs.clear()
        self.pair_headers.clear()
        self._pair_language_inputs.clear()

        for pair_data in project_data.get("language_pairs", []):
            pair_key = pair_data["pair_name"]
            header_title = pair_data.get("header_title", pair_key)
            widget = LanguagePairWidget(
                pair_key, self.currency_symbol, self.get_current_currency_code()
            )
            widget.remove_requested.connect(
                lambda w=widget: self._on_widget_remove_requested(w)
            )
            widget.subtotal_changed.connect(self.update_total)
            widget.name_changed.connect(
                lambda new_name, w=widget: self.on_pair_name_changed(w, new_name)
            )
            self.language_pairs[pair_key] = widget
            self.pairs_layout.insertWidget(self.pairs_layout.count() - 1, widget)
            self.pair_headers[pair_key] = header_title

            left_raw, right_raw = self._extract_pair_parts(pair_key)
            self._store_pair_language_inputs(
                pair_key,
                {"text": left_raw, "dict": False, "key": left_raw},
                {"text": right_raw, "dict": False, "key": right_raw},
            )

            services = pair_data.get("services", {})
            if "translation" in services:
                widget.translation_group.setChecked(True)
                widget.load_table_data(services["translation"])

            widget.set_report_sources(pair_data.get("report_sources", []), replace=True)
            pair_mode = pair_data.get("only_new_repeats")
            if pair_mode is None:
                pair_mode = self.only_new_repeats_mode
            widget.set_only_new_and_repeats_mode(bool(pair_mode))
            widget.set_discount_percent(pair_data.get("discount_percent", 0.0))
            widget.set_markup_percent(pair_data.get("markup_percent", 0.0))

        self._update_language_variant_regions_from_pairs(self.language_pairs.keys())

        self._update_language_names("ru" if self.lang_display_ru else "en")

        ps_rows = project_data.get("project_setup")
        ps_enabled = project_data.get("project_setup_enabled")
        ps_discount = project_data.get("project_setup_discount_percent", 0.0)
        ps_markup = project_data.get("project_setup_markup_percent", 0.0)
        if ps_rows is not None or ps_enabled is not None:
            if not self.project_setup_widget:
                self.project_setup_widget = ProjectSetupWidget(
                    self.project_setup_fee_spin.value(),
                    self.currency_symbol,
                    self.get_current_currency_code(),
                )
                self.project_setup_widget.remove_requested.connect(
                    self.remove_project_setup_widget
                )
                self.project_setup_widget.table.itemChanged.connect(
                    self.on_project_setup_item_changed
                )
                self.pairs_layout.insertWidget(0, self.project_setup_widget)
            if isinstance(ps_rows, dict):
                rows_to_load = ps_rows.get("rows", [])
                ps_discount = ps_rows.get("discount_percent", ps_discount)
                ps_markup = ps_rows.get("markup_percent", ps_markup)
            elif isinstance(ps_rows, list):
                rows_to_load = ps_rows
            elif ps_rows is None:
                rows_to_load = []
            else:
                rows_to_load = list(ps_rows) if ps_rows else []
            enabled_flag = True if ps_enabled is None else bool(ps_enabled)
            self.project_setup_widget.load_data(rows_to_load, enabled=enabled_flag)
            self.project_setup_widget.set_discount_percent(ps_discount)
            self.project_setup_widget.set_markup_percent(ps_markup)
            fee_value = project_data.get("project_setup_fee")
            if rows_to_load:
                first_vol = rows_to_load[0].get("volume")
                if isinstance(first_vol, (int, float)):
                    fee_value = first_vol
            if isinstance(fee_value, (int, float)):
                self.project_setup_fee_spin.blockSignals(True)
                self.project_setup_fee_spin.setValue(fee_value)
                self.project_setup_fee_spin.blockSignals(False)
                if self.project_setup_widget:
                    self.project_setup_widget.set_volume(fee_value)
        else:
            fee = project_data.get("project_setup_fee")
            if isinstance(fee, (int, float)):
                self.project_setup_fee_spin.blockSignals(True)
                self.project_setup_fee_spin.setValue(fee)
                self.project_setup_fee_spin.blockSignals(False)
                if self.project_setup_widget:
                    self.project_setup_widget.set_volume(fee)

        additional = project_data.get("additional_services")
        if additional is not None:
            self.additional_services_widget.load_data(additional)

        self.update_total()
