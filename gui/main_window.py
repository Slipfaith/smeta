import os
import shutil
import subprocess
import sys
import tempfile
import re
import traceback
from datetime import datetime
from typing import Dict, List, Any, Tuple, Optional

from PySide6.QtWidgets import (
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QGroupBox,
    QTextEdit,
    QFileDialog,
    QMessageBox,
    QScrollArea,
    QTabWidget,
    QSplitter,
    QComboBox,
    QSlider,
    QDoubleSpinBox,
    QInputDialog,
    QApplication,
)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QAction, QActionGroup

from logic.progress import Progress
from updater import APP_VERSION, AUTHOR, RELEASE_DATE, check_for_updates
from gui.language_pair import LanguagePairWidget
from gui.additional_services import AdditionalServicesWidget
from gui.project_manager_dialog import ProjectManagerDialog
from gui.project_setup_widget import ProjectSetupWidget
from gui.styles import APP_STYLE
from gui.utils import format_amount, format_language_display
from gui.rates_import_dialog import ExcelRatesDialog
from logic.excel_exporter import ExcelExporter
from logic.pdf_exporter import xlsx_to_pdf
from logic.user_config import load_languages, add_language
from logic.trados_xml_parser import parse_reports
from logic.service_config import ServiceConfig
from logic.pm_store import load_pm_history, save_pm_history
from logic.legal_entities import get_legal_entity_metadata, load_legal_entities
from logic.translation_config import tr
from logic.xml_parser_common import resolve_language_display
from logic.project_io import (
    save_project as save_project_file,
    load_project as load_project_file,
)
from logic.outlook_import import (
    OutlookMsgError,
    map_message_to_project_info,
    parse_msg_file,
)

CURRENCY_SYMBOLS = {"RUB": "₽", "EUR": "€", "USD": "$"}

class DropArea(QScrollArea):
    def __init__(self, callback, parent=None):
        super().__init__(parent)
        self._callback = callback
        self.setAcceptDrops(True)
        self.setWidgetResizable(True)

        self._base_style = """
            QScrollArea {
                border: 2px dashed #e5e7eb;
                border-radius: 8px;
                background-color: #fafafa;
            }
            QScrollArea[dragOver="true"] {
                border: 2px dashed #2563eb;
                background-color: #eff6ff;
            }
        """
        self.setStyleSheet(self._base_style)

    def disable_hint_style(self):
        self.setStyleSheet(
            """
            QScrollArea[dragOver="true"] {
                border: 2px dashed #2563eb;
                background-color: #eff6ff;
            }
        """
        )

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            all_paths = []
            xml_paths = []
            for url in urls:
                path = url.toLocalFile()
                all_paths.append(path)
                if path.lower().endswith(".xml") or path.lower().endswith(".XML"):
                    xml_paths.append(path)
            if xml_paths:
                event.acceptProposedAction()
                self.setProperty("dragOver", True)
                self.style().unpolish(self)
                self.style().polish(self)
                return
        event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.setProperty("dragOver", False)
        self.style().unpolish(self)
        self.style().polish(self)

    def dropEvent(self, event):
        self.setProperty("dragOver", False)
        self.style().unpolish(self)
        self.style().polish(self)

        if not event.mimeData().hasUrls():
            event.ignore()
            return

        urls = event.mimeData().urls()

        all_paths = []
        xml_paths = []

        for url in urls:
            path = url.toLocalFile()
            all_paths.append(path)

            try:
                if not os.path.exists(path) or not os.path.isfile(path):
                    continue
            except Exception:
                continue

            if path.lower().endswith((".xml", ".XML")):
                xml_paths.append(path)
            else:
                try:
                    with open(path, "r", encoding="utf-8", errors="ignore") as f:
                        first_line = f.readline().strip()
                        if first_line.startswith("<?xml") or "<" in first_line:
                            xml_paths.append(path)
                except Exception:
                    pass

        if xml_paths:
            try:
                self._callback(xml_paths)
                event.acceptProposedAction()
            except Exception as e:
                QMessageBox.critical(
                    None, "Ошибка", f"Ошибка при обработке файлов: {e}"
                )
        else:
            if all_paths:
                QMessageBox.warning(
                    None,
                    "Предупреждение",
                    f"Среди {len(all_paths)} перетащенных файлов не найдено ни одного XML файла.\n",
                    "Поддерживаются только файлы с расширением .xml",
                )
            event.ignore()


class ProjectInfoDropArea(QGroupBox):
    def __init__(self, title: str, callback, parent=None):
        super().__init__(title, parent)
        self._callback = callback
        self.setAcceptDrops(True)

    def _set_drag_state(self, active: bool):
        self.setProperty("dragOver", active)
        self.style().unpolish(self)
        self.style().polish(self)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                path = url.toLocalFile()
                if path.lower().endswith(".msg"):
                    event.acceptProposedAction()
                    self._set_drag_state(True)
                    return
        event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self._set_drag_state(False)

    def dropEvent(self, event):
        self._set_drag_state(False)
        if not event.mimeData().hasUrls():
            event.ignore()
            return

        msg_paths: List[str] = []
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(".msg"):
                msg_paths.append(path)

        if not msg_paths:
            event.ignore()
            return

        try:
            self._callback(msg_paths)
            event.acceptProposedAction()
        except Exception as exc:
            traceback.print_exc()
            QMessageBox.critical(
                self,
                "Ошибка обработки Outlook файла",
                str(exc) or "Не удалось обработать перетащенный .msg файл.",
            )
            event.ignore()


class TranslationCostCalculator(QMainWindow):
    def __init__(self):
        super().__init__()
        self.language_pairs: Dict[str, LanguagePairWidget] = {}
        self.pair_headers: Dict[str, str] = {}
        self._pair_language_inputs: Dict[str, Dict[str, Dict[str, Any]]] = {}
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
        self.excel_dialog = None
        self._import_pair_map: Dict[Tuple[str, str], str] = {}
        # Create labels early so slots triggered during initialization
        # (e.g. vat spin value changes) can safely update them.
        self.total_label = QLabel()
        self.discount_total_label = QLabel()
        self.markup_total_label = QLabel()
        self.setup_ui()
        self.setup_style()
        QTimer.singleShot(0, self.auto_check_for_updates)

    def setup_ui(self):
        self.setGeometry(100, 100, 1000, 600)
        self.setMinimumSize(600, 400)
        self.resize(1000, 650)
        self.update_title()

        lang = self.gui_lang

        self.project_menu = self.menuBar().addMenu(tr("Проект", lang))
        self.save_action = QAction(tr("Сохранить проект", lang), self)
        self.save_action.triggered.connect(self.save_project)
        self.project_menu.addAction(self.save_action)
        self.load_action = QAction(tr("Загрузить проект", lang), self)
        self.load_action.triggered.connect(self.load_project)
        self.project_menu.addAction(self.load_action)
        self.clear_action = QAction(tr("Очистить", lang), self)
        self.clear_action.triggered.connect(self.clear_all_data)
        self.project_menu.addAction(self.clear_action)

        self.export_menu = self.menuBar().addMenu(tr("Экспорт", lang))
        self.save_excel_action = QAction(tr("Сохранить Excel", lang), self)
        self.save_excel_action.triggered.connect(self.save_excel)
        self.export_menu.addAction(self.save_excel_action)
        self.save_pdf_action = QAction(tr("Сохранить PDF", lang), self)
        self.save_pdf_action.triggered.connect(self.save_pdf)
        self.export_menu.addAction(self.save_pdf_action)

        self.rates_menu = self.menuBar().addMenu(tr("Импорт ставок", lang))
        self.import_rates_action = QAction(tr("Импортировать из Excel", lang), self)
        self.import_rates_action.triggered.connect(self.import_rates_from_excel)
        self.rates_menu.addAction(self.import_rates_action)

        self.pm_action = QAction(tr("Проджект менеджер", lang), self)
        self.pm_action.triggered.connect(self.show_pm_dialog)
        self.menuBar().addAction(self.pm_action)

        self.update_menu = self.menuBar().addMenu(tr("Обновление", lang))
        self.check_updates_action = QAction(tr("Проверить обновления", lang), self)
        self.check_updates_action.triggered.connect(self.manual_update_check)
        self.update_menu.addAction(self.check_updates_action)

        self.about_action = QAction(tr("О программе", lang), self)
        self.about_action.triggered.connect(self.show_about_dialog)
        self.menuBar().addAction(self.about_action)

        self.language_menu = self.menuBar().addMenu(tr("Язык", lang))
        self.lang_action_group = QActionGroup(self)
        self.lang_ru_action = QAction(tr("Русский", lang), self)
        self.lang_en_action = QAction(tr("Английский", lang), self)
        self.lang_ru_action.setCheckable(True)
        self.lang_en_action.setCheckable(True)
        self.lang_action_group.addAction(self.lang_ru_action)
        self.lang_action_group.addAction(self.lang_en_action)
        self.language_menu.addAction(self.lang_ru_action)
        self.language_menu.addAction(self.lang_en_action)
        self.lang_ru_action.setChecked(self.gui_lang == "ru")
        self.lang_en_action.setChecked(self.gui_lang != "ru")
        self.lang_ru_action.triggered.connect(lambda: self.set_app_language("ru"))
        self.lang_en_action.triggered.connect(lambda: self.set_app_language("en"))
        self.update_menu_texts()

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QHBoxLayout()
        splitter = QSplitter(Qt.Horizontal)

        left_panel = self.create_left_panel()
        right_panel = self.create_right_panel()

        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)

        splitter.setSizes([600, 960])

        main_layout.addWidget(splitter)
        central_widget.setLayout(main_layout)

    def create_left_panel(self) -> QWidget:
        container = QWidget()
        lay = QVBoxLayout()
        lay.setSpacing(12)

        lang = self.gui_lang

        self.project_group = ProjectInfoDropArea(
            tr("Информация о проекте", lang), self.handle_project_info_drop
        )
        p = QVBoxLayout()
        p.setSpacing(8)
        self.project_name_label = QLabel(tr("Название проекта", lang) + ":")
        p.addWidget(self.project_name_label)
        self.project_name_edit = QLineEdit()
        p.addWidget(self.project_name_edit)
        self.client_name_label = QLabel(tr("Название клиента", lang) + ":")
        p.addWidget(self.client_name_label)
        self.client_name_edit = QLineEdit()
        p.addWidget(self.client_name_edit)
        self.contact_person_label = QLabel(tr("Контактное лицо", lang) + ":")
        p.addWidget(self.contact_person_label)
        self.contact_person_edit = QLineEdit()
        p.addWidget(self.contact_person_edit)
        self.email_label = QLabel(tr("Email", lang) + ":")
        p.addWidget(self.email_label)
        self.email_edit = QLineEdit()
        p.addWidget(self.email_edit)
        self.legal_entity_label = QLabel(tr("Юрлицо", lang) + ":")
        p.addWidget(self.legal_entity_label)
        self.legal_entity_combo = QComboBox()
        # Placeholder that indicates no legal entity selected yet
        self.legal_entity_placeholder = tr("Выберите юрлицо", lang)
        self.legal_entity_combo.addItem(self.legal_entity_placeholder)
        self.legal_entity_combo.addItems(self.legal_entities.keys())
        self.legal_entity_combo.setCurrentIndex(0)
        self.legal_entity_combo.currentTextChanged.connect(self.on_legal_entity_changed)
        p.addWidget(self.legal_entity_combo)
        self.currency_label = QLabel(tr("Валюта", lang) + ":")
        p.addWidget(self.currency_label)
        self.currency_combo = QComboBox()
        self.currency_placeholder = tr("Выберите валюту", lang)
        self.currency_combo.addItem(self.currency_placeholder)
        self.currency_combo.addItems(["RUB", "EUR", "USD"])
        self.currency_combo.setCurrentIndex(0)
        self.currency_combo.currentIndexChanged.connect(self.on_currency_index_changed)
        p.addWidget(self.currency_combo)
        self.convert_btn = QPushButton(tr("Конвертировать в рубли", lang))
        self.convert_btn.clicked.connect(self.convert_to_rub)
        p.addWidget(self.convert_btn)

        self.vat_label = QLabel(tr("НДС, %", lang) + ":")
        self.vat_spin = QDoubleSpinBox()
        self.vat_spin.setDecimals(2)
        self.vat_spin.setRange(0, 100)
        self.vat_spin.setValue(20.0)
        self.vat_spin.valueChanged.connect(self.update_total)
        self.vat_spin.wheelEvent = lambda event: event.ignore()

        vat_layout = QHBoxLayout()
        vat_layout.addWidget(self.vat_label)
        vat_layout.addWidget(self.vat_spin)
        p.addLayout(vat_layout)
        self.project_group.setLayout(p)
        lay.addWidget(self.project_group)
        # Initial state: no legal entity selected
        self.on_legal_entity_changed("")
        self.on_currency_changed(self.get_current_currency_code())

        self.pairs_group = QGroupBox(tr("Языковые пары", lang))
        pg = QVBoxLayout()
        pg.setSpacing(8)

        mode = QHBoxLayout()
        self.language_names_label = QLabel(tr("Названия языков", lang) + ":")
        mode.addWidget(self.language_names_label)
        mode.addStretch(1)
        mode.addWidget(QLabel("EN"))
        self.lang_mode_slider = QSlider(Qt.Horizontal)
        self.lang_mode_slider.setRange(0, 1)
        self.lang_mode_slider.setValue(1)
        self.lang_mode_slider.setFixedWidth(70)
        self.lang_mode_slider.valueChanged.connect(self.on_lang_mode_changed)
        mode.addWidget(self.lang_mode_slider)
        mode.addWidget(QLabel("RU"))
        pg.addLayout(mode)

        add_pair = QHBoxLayout()
        self.source_lang_combo = self._make_lang_combo()
        self.source_lang_combo.setEditable(True)
        add_pair.addWidget(self.source_lang_combo)
        add_pair.addWidget(QLabel("→"))
        self.target_lang_combo = self._make_lang_combo()
        self.target_lang_combo.setEditable(True)
        add_pair.addWidget(self.target_lang_combo)
        pg.addLayout(add_pair)

        self.add_pair_btn = QPushButton(tr("Добавить языковую пару", lang))
        self.add_pair_btn.clicked.connect(self.add_language_pair)
        pg.addWidget(self.add_pair_btn)

        self.current_pairs_label = QLabel(tr("Текущие пары", lang) + ":")
        pg.addWidget(self.current_pairs_label)
        self.pairs_list = QTextEdit()
        self.pairs_list.setMaximumHeight(100)
        self.pairs_list.setReadOnly(True)
        pg.addWidget(self.pairs_list)

        info_layout = QHBoxLayout()
        self.language_pairs_count_label = QLabel(f"{tr('Загружено языковых пар', lang)}: 0")
        info_layout.addWidget(self.language_pairs_count_label)
        info_layout.addStretch()
        self.clear_pairs_btn = QPushButton(tr("Очистить", lang))
        self.clear_pairs_btn.clicked.connect(self.clear_language_pairs)
        info_layout.addWidget(self.clear_pairs_btn)
        pg.addLayout(info_layout)

        setup_layout = QHBoxLayout()
        self.project_setup_label = QLabel(
            tr("Запуск и управление проектом", lang) + ":",
        )
        setup_layout.addWidget(self.project_setup_label)
        self.project_setup_fee_spin = QDoubleSpinBox()
        self.project_setup_fee_spin.setDecimals(2)
        self.project_setup_fee_spin.setSingleStep(0.25)
        self.project_setup_fee_spin.setMinimum(0.5)
        self.project_setup_fee_spin.setValue(0.5)
        setup_layout.addWidget(self.project_setup_fee_spin)
        setup_layout.addStretch()
        pg.addLayout(setup_layout)

        self.add_lang_group = QGroupBox(tr("Добавить язык в справочник", lang))
        lg = QVBoxLayout()
        lg.setSpacing(8)
        r1 = QHBoxLayout()
        self.lang_ru_label = QLabel(tr("Название RU", lang) + ":")
        r1.addWidget(self.lang_ru_label)
        self.new_lang_ru = QLineEdit()
        self.new_lang_ru.setPlaceholderText(tr("Валирийский", "ru"))
        r1.addWidget(self.new_lang_ru)
        lg.addLayout(r1)
        r2 = QHBoxLayout()
        self.lang_en_label = QLabel(tr("Название EN", lang) + ":")
        r2.addWidget(self.lang_en_label)
        self.new_lang_en = QLineEdit()
        self.new_lang_en.setPlaceholderText(tr("Valyrian", "en"))
        r2.addWidget(self.new_lang_en)
        lg.addLayout(r2)
        self.btn_add_lang = QPushButton(tr("Добавить язык", lang))
        self.btn_add_lang.clicked.connect(self.handle_add_language)
        lg.addWidget(self.btn_add_lang)
        self.add_lang_group.setLayout(lg)
        pg.addWidget(self.add_lang_group)

        self.pairs_group.setLayout(pg)
        lay.addWidget(self.pairs_group)

        lay.addStretch()
        container.setLayout(lay)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(container)
        scroll.setMinimumWidth(280)
        return scroll

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
        self.update_menu_texts()

    def update_menu_texts(self):
        lang = self.gui_lang
        self.project_menu.setTitle(tr("Проект", lang))
        self.save_action.setText(tr("Сохранить проект", lang))
        self.load_action.setText(tr("Загрузить проект", lang))
        self.clear_action.setText(tr("Очистить", lang))
        self.export_menu.setTitle(tr("Экспорт", lang))
        self.save_excel_action.setText(tr("Сохранить Excel", lang))
        self.save_pdf_action.setText(tr("Сохранить PDF", lang))
        self.rates_menu.setTitle(tr("Импорт ставок", lang))
        self.import_rates_action.setText(tr("Импортировать из Excel", lang))
        self.pm_action.setText(tr("Проджект менеджер", lang))
        self.update_menu.setTitle(tr("Обновление", lang))
        self.check_updates_action.setText(tr("Проверить обновления", lang))
        self.about_action.setText(tr("О программе", lang))
        self.language_menu.setTitle(tr("Язык", lang))
        self.lang_ru_action.setText(tr("Русский", lang))
        self.lang_en_action.setText(tr("Английский", lang))

    def manual_update_check(self):
        check_for_updates(self, force=True)

    def auto_check_for_updates(self):
        check_for_updates(self, force=False)

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
        if code:
            normalized = str(code).strip().upper()
            idx = self.currency_combo.findText(normalized, Qt.MatchFixedString)
            if idx < 0:
                for i in range(1, self.currency_combo.count()):
                    text = self.currency_combo.itemText(i).strip().upper()
                    if text == normalized:
                        idx = i
                        break
            if idx >= 0:
                self.currency_combo.setCurrentIndex(idx)
                return True
        self.currency_combo.setCurrentIndex(0)
        return False

    def on_currency_changed(self, code: str):
        self.currency_symbol = CURRENCY_SYMBOLS.get(code, code)
        if getattr(self, "project_setup_widget", None):
            self.project_setup_widget.set_currency(self.currency_symbol, code)
        for w in self.language_pairs.values():
            w.set_currency(self.currency_symbol, code)
        if getattr(self, "additional_services_widget", None):
            self.additional_services_widget.set_currency(self.currency_symbol, code)
        if getattr(self, "convert_btn", None):
            self.convert_btn.setEnabled(code == "USD")

    def get_selected_legal_entity(self) -> str:
        """Return currently selected legal entity or empty string if none."""
        idx = self.legal_entity_combo.currentIndex()
        if idx <= 0:
            return ""
        return self.legal_entity_combo.currentText()

    def convert_to_rub(self):
        """Convert all rates from USD to RUB using user-provided rate."""
        if self.get_current_currency_code() != "USD":
            return
        lang = self.gui_lang
        rate, ok = QInputDialog.getDouble(
            self,
            tr("Курс USD", lang),
            tr("1 USD в рублях", lang),
            0.0,
            0.0,
            1000000.0,
            4,
        )
        if not ok or rate <= 0:
            return
        if getattr(self, "project_setup_widget", None):
            self.project_setup_widget.convert_rates(rate)
        for w in self.language_pairs.values():
            w.convert_rates(rate)
        if getattr(self, "additional_services_widget", None):
            self.additional_services_widget.convert_rates(rate)
        self.set_currency_code("RUB")
        self.update_total()

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

    def create_right_panel(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout()
        self.tabs = QTabWidget()
        gui_lang = self.gui_lang
        est_lang = "ru" if self.lang_display_ru else "en"

        self.pairs_scroll = QScrollArea()
        self.pairs_scroll.setWidgetResizable(True)
        self.pairs_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.pairs_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        self.pairs_container_widget = QWidget()
        self.pairs_layout = QVBoxLayout()
        self.pairs_layout.setSpacing(12)

        self.only_new_repeats_btn = QPushButton(
            tr("Только новые слова и повторы", gui_lang)
        )
        self.only_new_repeats_btn.clicked.connect(self.toggle_only_new_repeats_mode)
        self.pairs_layout.addWidget(self.only_new_repeats_btn)

        self.project_setup_widget = ProjectSetupWidget(
            self.project_setup_fee_spin.value(),
            self.currency_symbol,
            self.get_current_currency_code(),
            lang=est_lang,
        )
        self.project_setup_widget.remove_requested.connect(
            self.remove_project_setup_widget
        )
        self.project_setup_widget.subtotal_changed.connect(self.update_total)
        self.pairs_layout.addWidget(self.project_setup_widget)
        self.project_setup_fee_spin.valueChanged.connect(
            self.update_project_setup_volume_from_spin
        )
        self.project_setup_widget.table.itemChanged.connect(
            self.on_project_setup_item_changed
        )

        self.drop_hint_label = QLabel(
            tr(
                "Перетащите XML файлы отчетов Trados или Smartcat сюда для автоматического заполнения",
                gui_lang,
            )
        )
        self.drop_hint_label.setStyleSheet(
            """
            QLabel {
                color: #9ca3af;
                font-style: italic;
                padding: 24px;
                text-align: center;
                background-color: #f9fafb;
                border: 2px dashed #e5e7eb;
                border-radius: 8px;
                margin: 16px 0;
            }
        """
        )
        self.drop_hint_label.setAlignment(Qt.AlignCenter)
        self.pairs_layout.addWidget(self.drop_hint_label)

        self.pairs_layout.addStretch()

        self.pairs_container_widget.setLayout(self.pairs_layout)
        self.pairs_scroll.setWidget(self.pairs_container_widget)

        self.pairs_scroll.setAcceptDrops(True)
        self.setup_drag_drop()

        self.tabs.addTab(self.pairs_scroll, tr("Языковые пары", gui_lang))

        self.additional_services_widget = AdditionalServicesWidget(
            self.currency_symbol,
            self.get_current_currency_code(),
            lang=est_lang,
        )
        self.additional_services_widget.subtotal_changed.connect(self.update_total)
        add_scroll = QScrollArea()
        add_scroll.setWidget(self.additional_services_widget)
        add_scroll.setWidgetResizable(True)
        self.tabs.addTab(add_scroll, tr("Дополнительные услуги", gui_lang))

        lay.addWidget(self.tabs)

        # ``total_label`` and adjustment labels are created in ``__init__`` so that
        # early signal emissions can reference them without raising errors. Here we
        # simply style them and add them to the layout.
        self.markup_total_label.setAlignment(Qt.AlignRight)
        self.markup_total_label.setStyleSheet(
            "font-size: 12px; padding: 4px; color: #555;"
        )
        self.markup_total_label.hide()
        lay.addWidget(self.markup_total_label)
        self.discount_total_label.setAlignment(Qt.AlignRight)
        self.discount_total_label.setStyleSheet(
            "font-size: 12px; padding: 4px; color: #555;"
        )
        self.discount_total_label.hide()
        lay.addWidget(self.discount_total_label)
        self.total_label.setAlignment(Qt.AlignRight)
        self.total_label.setStyleSheet(
            "font-weight: bold; font-size: 14px; padding: 6px; color: #333;"
        )
        lay.addWidget(self.total_label)

        w.setLayout(lay)
        self.update_total()
        return w

    def setup_drag_drop(self):
        drop_area = DropArea(self.handle_xml_drop)

        drop_area.setWidget(self.pairs_container_widget)

        self.tabs.removeTab(0)
        lang = self.gui_lang
        self.tabs.insertTab(0, drop_area, tr("Языковые пары", lang))
        self.pairs_scroll = drop_area

    def handle_project_info_drop(self, paths: List[str]):
        lang = self.gui_lang
        errors: List[str] = []

        for path in paths:
            try:
                message = parse_msg_file(path)
                result = map_message_to_project_info(message)
                self._apply_project_info_result(result, path)
                return
            except (OutlookMsgError, RuntimeError) as exc:
                errors.append(f"{os.path.basename(path)}: {exc}")

        if errors:
            QMessageBox.warning(
                self,
                tr("Ошибка обработки Outlook файла", lang),
                "\n".join(errors),
            )

    def _apply_project_info_result(self, result, source_path: str):
        lang = self.gui_lang
        data = result.data

        updated_fields: List[str] = []
        manual_checks: List[str] = list(result.warnings)

        def update_field(widget, value: Optional[str], label_key: str):
            if value:
                widget.setText(value)
                updated_fields.append(tr(label_key, lang))
            else:
                widget.clear()

        update_field(self.project_name_edit, data.project_name, "Название проекта")
        update_field(self.client_name_edit, data.client_name, "Название клиента")
        update_field(self.contact_person_edit, data.contact_name, "Контактное лицо")
        update_field(self.email_edit, data.email, "Email")

        legal_entity_value = (data.legal_entity or "").strip()
        if legal_entity_value.lower() == "logrus it usa":
            legal_entity_value = "Logrus IT"
            data.legal_entity = legal_entity_value
            if self.lang_mode_slider.value() != 0:
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

        if data.currency_code:
            if self.set_currency_code(data.currency_code):
                updated_fields.append(tr("Валюта", lang))
            else:
                self.set_currency_code(None)
                manual_checks.append(f"{tr('Валюта', lang)}: {data.currency_code}")
        else:
            self.set_currency_code(None)

        missing = [tr(name, lang) for name in result.missing_fields]

        sender_parts: List[str] = []
        if result.sender_name:
            sender_parts.append(result.sender_name)
        if result.sender_email:
            sender_parts.append(result.sender_email)

        def format_section(title: str, values: List[str]) -> Optional[str]:
            unique_values = list(dict.fromkeys(v for v in values if v))
            if not unique_values:
                return None
            bullets = "\n  • ".join(unique_values)
            return f"{title}:\n  • {bullets}"

        message_sections: List[str] = [
            f"{tr('Outlook письмо', lang)}: {os.path.basename(source_path)}"
        ]

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

        if result.sent_at:
            message_sections.append(f"{tr('Дата отправки', lang)}: {result.sent_at}")

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
        name = self.current_pm.get("name_ru") or self.current_pm.get("name_en") or ""
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
        total = 0.0
        discount_total = 0.0
        markup_total = 0.0
        if getattr(self, "project_setup_widget", None):
            total += self.project_setup_widget.get_subtotal()
            discount_total += self.project_setup_widget.get_discount_amount()
            markup_total += self.project_setup_widget.get_markup_amount()
        for w in self.language_pairs.values():
            total += w.get_subtotal()
            discount_total += w.get_discount_amount()
            markup_total += w.get_markup_amount()
        if getattr(self, "additional_services_widget", None):
            total += self.additional_services_widget.get_subtotal()
            discount_total += self.additional_services_widget.get_discount_amount()
            markup_total += self.additional_services_widget.get_markup_amount()
        lang = "ru" if self.lang_display_ru else "en"
        vat_rate = (
            self.vat_spin.value() / 100 if self.vat_spin.isEnabled() else 0.0
        )

        symbol_suffix = f" {self.currency_symbol}" if self.currency_symbol else ""
        if markup_total > 0:
            self.markup_total_label.setText(
                f"{tr('Сумма наценки', lang)}: {format_amount(markup_total, lang)}{symbol_suffix}"
            )
            self.markup_total_label.show()
        else:
            self.markup_total_label.hide()
        if discount_total > 0:
            self.discount_total_label.setText(
                f"{tr('Сумма скидки', lang)}: {format_amount(discount_total, lang)}{symbol_suffix}"
            )
            self.discount_total_label.show()
        else:
            self.discount_total_label.hide()

        if vat_rate > 0:
            vat_amount = total * vat_rate
            total_with_vat = total + vat_amount
            self.total_label.setText(
                f"{tr('Итого', lang)}: {format_amount(total_with_vat, lang)}{symbol_suffix} {tr('с НДС', lang)}. "
                f"{tr('НДС', lang)}: {format_amount(vat_amount, lang)}{symbol_suffix}"
            )
        else:
            self.total_label.setText(
                f"{tr('Итого', lang)}: {format_amount(total, lang)}{symbol_suffix}"
            )

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

    def _labels_from_entry(self, entry: Dict[str, Any]) -> Dict[str, str]:
        if not entry:
            return {"en": "", "ru": ""}

        key_value = entry.get("key", "").strip()
        text = entry.get("text", "").strip()
        en_value = entry.get("en", "").strip()
        ru_value = entry.get("ru", "").strip()

        resolved_en = ""
        resolved_ru = ""
        if key_value:
            resolved_en = resolve_language_display(key_value, locale="en") or ""
            resolved_ru = resolve_language_display(key_value, locale="ru") or ""
        if not resolved_en and text:
            resolved_en = resolve_language_display(text, locale="en") or ""
        if not resolved_ru and text:
            resolved_ru = resolve_language_display(text, locale="ru") or ""
        if not resolved_en and en_value:
            resolved_en = resolve_language_display(en_value, locale="en") or ""
        if not resolved_ru and ru_value:
            resolved_ru = resolve_language_display(ru_value, locale="ru") or ""

        en_name = resolved_en or en_value or ru_value or text or key_value
        ru_name = resolved_ru or ru_value or en_value or text or key_value

        return {
            "en": self._prepare_language_label(en_name, "en"),
            "ru": self._prepare_language_label(ru_name, "ru"),
        }

    def add_language_pair(self):
        src = self._parse_combo(self.source_lang_combo)
        tgt = self._parse_combo(self.target_lang_combo)
        if not src["text"] or not tgt["text"]:
            lang = self.gui_lang
            QMessageBox.warning(
                self, tr("Ошибка", lang), tr("Выберите/введите оба языка", lang)
            )
            return

        def key_name(obj: Dict[str, Any]) -> str:
            return obj["en"] or obj["ru"] or obj["text"]

        left_key = key_name(src)
        right_key = key_name(tgt)
        pair_key = f"{left_key} → {right_key}"

        if pair_key in self.language_pairs:
            lang = self.gui_lang
            QMessageBox.warning(
                self, tr("Ошибка", lang), tr("Такая языковая пара уже существует", lang)
            )
            return

        self._store_pair_language_inputs(pair_key, src, tgt, left_key, right_key)

        labels = self._pair_language_inputs.get(pair_key, {})
        src_labels = self._labels_from_entry(labels.get("source", {}))
        tgt_labels = self._labels_from_entry(labels.get("target", {}))
        lang_key = "ru" if self.lang_display_ru else "en"
        locale = "ru" if self.lang_display_ru else "en"
        src_value = src_labels[lang_key]
        tgt_value = tgt_labels[lang_key]
        if not src_value:
            src_value = self._prepare_language_label(src.get("text", ""), locale)
        if not tgt_value:
            tgt_value = self._prepare_language_label(tgt.get("text", ""), locale)
        display_name = f"{src_value} - {tgt_value}"
        header_value = tgt_labels[lang_key]
        if not header_value:
            locale = "ru" if self.lang_display_ru else "en"
            header_value = self._prepare_language_label(tgt.get("text", ""), locale)
        self.pair_headers[pair_key] = header_value

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
        if self.only_new_repeats_mode:
            widget.set_only_new_and_repeats_mode(True)

        self.pairs_layout.insertWidget(self.pairs_layout.count() - 1, widget)

        self.update_pairs_list()
        self.update_total()

        self.source_lang_combo.setCurrentIndex(0)
        self.target_lang_combo.setCurrentIndex(0)

    def _pair_sort_key(self, pair_key: str) -> str:
        _, right = self._extract_pair_parts(pair_key)
        return right or pair_key

    def _extract_pair_parts(self, pair_key: str) -> Tuple[str, str]:
        for sep in (" → ", " - "):
            if sep in pair_key:
                left, right = pair_key.split(sep, 1)
                return left.strip(), right.strip()
        return pair_key.strip(), ""

    def _prepare_language_label(self, name: str, locale: str) -> str:
        if not name:
            return ""
        return format_language_display(name, locale)

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

    def _display_pair_name(self, pair_key: str) -> str:
        lang = "ru" if self.lang_display_ru else "en"
        entries = self._pair_language_inputs.get(pair_key)
        if entries:
            left_labels = self._labels_from_entry(entries.get("source", {}))
            right_labels = self._labels_from_entry(entries.get("target", {}))
            if right_labels.get("en") or right_labels.get("ru"):
                return f"{left_labels[lang]} - {right_labels[lang]}"
            return left_labels[lang]

        left_key, right_key = self._extract_pair_parts(pair_key)
        if not right_key:
            return self._prepare_language_label(left_key, lang)
        left = self._find_language_by_key(left_key)
        right = self._find_language_by_key(right_key)
        return f"{left[lang]} - {right[lang]}"

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

    def remove_language_pair(self, pair_key: str):
        widget = self.language_pairs.pop(pair_key, None)
        if widget:
            widget.setParent(None)
            self.pair_headers.pop(pair_key, None)
        self._pair_language_inputs.pop(pair_key, None)
        self.update_pairs_list()
        self.update_total()

    def clear_language_pairs(self):
        for w in self.language_pairs.values():
            w.setParent(None)
        self.language_pairs.clear()
        self.pair_headers.clear()
        self._pair_language_inputs.clear()
        self.update_pairs_list()
        self.update_total()

    def clear_all_data(self):
        """Reset all user-entered and loaded data."""
        self.project_name_edit.clear()
        self.client_name_edit.clear()
        self.contact_person_edit.clear()
        self.email_edit.clear()
        self.legal_entity_combo.setCurrentIndex(0)
        self.currency_combo.setCurrentIndex(0)
        self.vat_spin.setValue(20.0)

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

        self.clear_language_pairs()
        if getattr(self, "additional_services_widget", None):
            self.additional_services_widget.load_data([])

        self.only_new_repeats_mode = False
        if getattr(self, "only_new_repeats_btn", None):
            self.only_new_repeats_btn.setText(
                tr("Только новые слова и повторы", self.gui_lang)
            )

        self.update_total()

    def import_rates_from_excel(self) -> None:
        if not self.language_pairs:
            QMessageBox.warning(self, "Ошибка", "Сначала добавьте языковые пары")
            return
        pairs = []
        pair_map = {}
        for key in self.language_pairs:
            parts = re.split(r"\s*(?:→|->|-|>)\s*", key, maxsplit=1)
            if len(parts) != 2:
                continue
            src, tgt = parts
            pairs.append((src, tgt))
            pair_map[(src, tgt)] = key
        self._import_pair_map = pair_map
        self.excel_dialog = ExcelRatesDialog(pairs, self)
        self.excel_dialog.finished.connect(self._on_rates_dialog_closed)
        self.excel_dialog.apply_requested.connect(self._apply_rates_from_dialog)
        self.excel_dialog.show()

    def _on_rates_dialog_closed(self, result: int) -> None:
        self._apply_rates_from_dialog()
        if self.excel_dialog:
            self.excel_dialog.deleteLater()
            self.excel_dialog = None
        self._import_pair_map = {}

    def _apply_rates_from_dialog(self) -> None:
        if not self.excel_dialog:
            return
        rate_key = self.excel_dialog.selected_rate_key()
        for match in self.excel_dialog.selected_rates():
            if not match.rates:
                continue
            pair_key = self._import_pair_map.get((match.gui_source, match.gui_target))
            if not pair_key:
                continue
            widget = self.language_pairs.get(pair_key)
            if widget:
                widget.set_basic_rate(match.rates.get(rate_key, 0))

    def handle_xml_drop(self, paths: List[str], replace: bool = False):
        try:
            data, warnings, report_sources = parse_reports(paths)

            if warnings:
                warning_msg = f"Предупреждения при обработке файлов:\n" + "\n".join(
                    warnings
                )
                QMessageBox.warning(self, "Предупреждение", warning_msg)

            if not data:
                QMessageBox.warning(
                    self,
                    "Результат обработки",
                    "В XML файлах не найдено данных о языковых парах.\n"
                    "Возможные причины:\n"
                    "1. XML файлы имеют нестандартную структуру\n"
                    "2. Не найдены элементы LanguageDirection\n"
                    "3. Отсутствуют данные о языках или объемах\n"
                    "Проверьте консоль для детальной информации.",
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

            result_msg = f"Обработка завершена!\n\n"
            if added_pairs > 0:
                result_msg += f"Добавлено новых языковых пар: {added_pairs}\n"
            if updated_pairs > 0:
                result_msg += f"Обновлено существующих пар: {updated_pairs}\n"
            result_msg += f"\nВсего обработано языковых пар: {len(data)}"

        except Exception as e:
            error_msg = f"Ошибка при обработке XML файлов: {str(e)}"
            QMessageBox.critical(self, "Ошибка", error_msg)

    def collect_project_data(self) -> Dict[str, Any]:
        data = {
            "project_name": self.project_name_edit.text(),
            "client_name": self.client_name_edit.text(),
            "contact_person": self.contact_person_edit.text(),
            "email": self.email_edit.text(),
            "legal_entity": self.get_selected_legal_entity(),
            "currency": self.get_current_currency_code(),
            "language_pairs": [],
            "additional_services": [],
            "pm_name": self.current_pm.get(
                "name_ru" if self.lang_display_ru else "name_en", ""
            ),
            "pm_email": self.current_pm.get("email", ""),
            "project_setup_fee": self.project_setup_fee_spin.value(),
            "project_setup_enabled": (
                self.project_setup_widget.is_enabled()
                if self.project_setup_widget
                else False
            ),
            "project_setup": (
                self.project_setup_widget.get_data()
                if self.project_setup_widget
                and self.project_setup_widget.is_enabled()
                else []
            ),
            "project_setup_discount_percent": (
                self.project_setup_widget.get_discount_percent()
                if self.project_setup_widget
                else 0.0
            ),
            "project_setup_markup_percent": (
                self.project_setup_widget.get_markup_percent()
                if self.project_setup_widget
                else 0.0
            ),
            "vat_rate": self.vat_spin.value() if self.vat_spin.isEnabled() else 0,
            "only_new_repeats_mode": self.only_new_repeats_mode,
        }
        total_discount_amount = 0.0
        total_markup_amount = 0.0
        project_setup_discount_amount = 0.0
        project_setup_markup_amount = 0.0
        if getattr(self, "project_setup_widget", None):
            project_setup_discount_amount = (
                self.project_setup_widget.get_discount_amount()
            )
            total_discount_amount += project_setup_discount_amount
            project_setup_markup_amount = (
                self.project_setup_widget.get_markup_amount()
            )
            total_markup_amount += project_setup_markup_amount
        data["project_setup_discount_amount"] = project_setup_discount_amount
        data["project_setup_markup_amount"] = project_setup_markup_amount
        for pair_key, pair_widget in self.language_pairs.items():
            p = pair_widget.get_data()
            if p["services"]:
                p["header_title"] = self.pair_headers.get(
                    pair_key, pair_widget.pair_name
                )
                data["language_pairs"].append(p)
                total_discount_amount += p.get("discount_amount", 0.0)
                total_markup_amount += p.get("markup_amount", 0.0)
        data["language_pairs"].sort(
            key=lambda x: (
                x.get("pair_name", "").split(" - ")[1]
                if " - " in x.get("pair_name", "")
                else x.get("pair_name", "")
            )
        )
        additional = self.additional_services_widget.get_data()
        if additional:
            data["additional_services"] = additional
            total_discount_amount += sum(
                block.get("discount_amount", 0.0) for block in additional
            )
            total_markup_amount += sum(
                block.get("markup_amount", 0.0) for block in additional
            )
        data["total_discount_amount"] = total_discount_amount
        data["total_markup_amount"] = total_markup_amount
        return data

    def save_excel(self):
        export_lang = "ru" if self.lang_display_ru else "en"
        if not self.get_selected_legal_entity():
            lang = self.gui_lang
            QMessageBox.warning(
                self,
                tr("Предупреждение", lang),
                tr("Выберите юрлицо", lang),
            )
            return
        if not self.client_name_edit.text().strip():
            QMessageBox.warning(self, "Ошибка", "Введите название клиента")
            return
        project_data = self.collect_project_data()
        if (
            not project_data["language_pairs"]
            and not project_data["additional_services"]
        ):
            QMessageBox.warning(self, "Ошибка", "Добавьте хотя бы одну услугу")
            return

        if any(r.get("rate", 0) == 0 for r in project_data.get("project_setup", [])):
            lang = self.gui_lang
            msg = tr("Ставка для \"{0}\" равна 0. Продолжить?", lang).format(
                tr("Запуск и управление проектом", lang)
            )
            reply = QMessageBox.question(
                self,
                tr("Предупреждение", lang),
                msg,
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply == QMessageBox.No:
                return

        client_name = project_data["client_name"].replace(" ", "_")
        entity_for_file = self.get_selected_legal_entity().replace(" ", "_")
        currency = self.get_current_currency_code()
        date_str = datetime.now().strftime("%Y-%m-%d")
        filename = f"{date_str}-{entity_for_file}-{currency}-{client_name}.xlsx"

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить Excel файл", filename, "Excel files (*.xlsx)"
        )
        if not file_path:
            return

        entity_name = self.get_selected_legal_entity()
        template_path = self.legal_entities.get(entity_name)
        exporter = ExcelExporter(
            template_path,
            currency=self.get_current_currency_code(),
            lang=export_lang,
        )
        with Progress(parent=self) as progress:
            success = exporter.export_to_excel(
                project_data, file_path, progress_callback=progress.on_progress
            )
        if success:
            self._show_file_saved_message(file_path)
        else:
            QMessageBox.critical(self, "Ошибка", "Не удалось сохранить файл")

    def save_pdf(self):
        export_lang = "ru" if self.lang_display_ru else "en"
        if not self.get_selected_legal_entity():
            lang = self.gui_lang
            QMessageBox.warning(
                self,
                tr("Предупреждение", lang),
                tr("Выберите юрлицо", lang),
            )
            return
        if not self.project_name_edit.text().strip():
            QMessageBox.warning(self, "Ошибка", "Введите название проекта")
            return
        if not self.client_name_edit.text().strip():
            QMessageBox.warning(self, "Ошибка", "Введите название клиента")
            return
        project_data = self.collect_project_data()
        if any(r.get("rate", 0) == 0 for r in project_data.get("project_setup", [])):
            lang = self.gui_lang
            msg = tr("Ставка для \"{0}\" равна 0. Продолжить?", lang).format(
                tr("Запуск и управление проектом", lang)
            )
            reply = QMessageBox.question(
                self,
                tr("Предупреждение", lang),
                msg,
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply == QMessageBox.No:
                return
        client_name = project_data["client_name"].replace(" ", "_")
        entity_for_file = self.get_selected_legal_entity().replace(" ", "_")
        currency = self.get_current_currency_code()
        date_str = datetime.now().strftime("%Y-%m-%d")
        filename = f"{date_str}-{entity_for_file}-{currency}-{client_name}.pdf"
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить PDF файл", filename, "PDF files (*.pdf)"
        )
        if not file_path:
            return
        template_path = self.legal_entities.get(self.get_selected_legal_entity())
        exporter = ExcelExporter(
            template_path,
            currency=currency,
            lang=export_lang,
        )
        with Progress(parent=self) as progress:
            def on_excel_progress(percent: int, message: str) -> None:
                progress.on_progress(int(percent * 0.8), message)

            try:
                with tempfile.TemporaryDirectory() as tmpdir:
                    xlsx_path = os.path.join(tmpdir, "quotation.xlsx")
                    pdf_path = os.path.join(tmpdir, "quotation.pdf")
                    if not exporter.export_to_excel(
                        project_data,
                        xlsx_path,
                        fit_to_page=True,
                        progress_callback=on_excel_progress,
                    ):
                        QMessageBox.critical(self, "Ошибка", "Не удалось подготовить файл")
                        return
                    progress.set_label("Конвертация в PDF")
                    progress.set_value(80)
                    if not xlsx_to_pdf(xlsx_path, pdf_path, lang=export_lang):
                        QMessageBox.critical(
                            self, "Ошибка", "Не удалось конвертировать в PDF"
                        )
                        return
                    progress.set_value(100)
                    shutil.copyfile(pdf_path, file_path)
                self._show_file_saved_message(file_path)
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить PDF: {e}")

    def save_project(self):
        if not self.project_name_edit.text().strip():
            QMessageBox.warning(self, "Ошибка", "Введите название проекта")
            return
        project_data = self.collect_project_data()
        project_name = project_data["project_name"].replace(" ", "_")
        filename = f"Проект_{project_name}.json"
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить проект", filename, "JSON files (*.json)"
        )
        if not file_path:
            return
        if save_project_file(project_data, file_path):
            QMessageBox.information(self, "Успех", f"Проект сохранен: {file_path}")
        else:
            QMessageBox.critical(self, "Ошибка", "Не удалось сохранить проект")

    def _show_file_saved_message(self, file_path: str) -> None:
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setWindowTitle("Успех")
        msg_box.setText(f"Файл сохранен:\n{file_path}")
        open_button = msg_box.addButton("Открыть папку", QMessageBox.ActionRole)
        msg_box.addButton(QMessageBox.Ok)
        msg_box.setDefaultButton(QMessageBox.Ok)
        msg_box.exec()
        if msg_box.clickedButton() == open_button:
            self._reveal_file_in_explorer(file_path)

    def _reveal_file_in_explorer(self, file_path: str) -> None:
        if not os.path.exists(file_path):
            QMessageBox.warning(self, "Ошибка", "Файл не найден для открытия в проводнике")
            return
        try:
            if sys.platform.startswith("win"):
                subprocess.run(
                    ["explorer", f"/select,{os.path.normpath(file_path)}"], check=False
                )
            elif sys.platform == "darwin":
                subprocess.run(["open", "-R", file_path], check=False)
            else:
                directory = os.path.dirname(file_path) or "."
                subprocess.run(["xdg-open", directory], check=False)
        except Exception as exc:
            QMessageBox.warning(
                self,
                "Ошибка",
                f"Не удалось открыть проводник:\n{exc}",
            )

    def load_project(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Загрузить проект", "", "JSON files (*.json)"
        )
        if not file_path:
            return
        project_data = load_project_file(file_path)
        if project_data is None:
            QMessageBox.critical(self, "Ошибка", "Не удалось загрузить проект")
            return
        self.load_project_data(project_data)
        QMessageBox.information(self, "Успех", "Проект загружен")

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
