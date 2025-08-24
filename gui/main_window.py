import json
import os
import shutil
import tempfile
import re
import sys
import threading
from datetime import datetime
from typing import Dict, List, Any

import langcodes
import pycountry

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
    QApplication,
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QAction

from logic.progress import Progress

from gui.language_pair import LanguagePairWidget
from gui.additional_services import AdditionalServicesWidget
from gui.project_manager_dialog import ProjectManagerDialog
from gui.project_setup_widget import ProjectSetupWidget
from gui.styles import APP_STYLE
from gui.utils import shorten_locale
from logic.excel_exporter import ExcelExporter
from logic.pdf_exporter import xlsx_to_pdf
from logic.user_config import load_languages, add_language
from logic.trados_xml_parser import parse_reports
from logic.service_config import ServiceConfig
from logic.pm_store import load_pm_history, save_pm_history
from logic.legal_entities import load_legal_entities
from logic.translation_config import tr
from logic.com_utils import get_excel_app, close_excel_app

CURRENCY_SYMBOLS = {"RUB": "₽", "EUR": "€", "USD": "$"}


class DropArea(QScrollArea):
    """QScrollArea, принимающая перетаскивание XML-файлов и отдающая пути в колбек."""

    def __init__(self, callback, parent=None):
        super().__init__(parent)
        self._callback = callback
        self.setAcceptDrops(True)
        self.setWidgetResizable(True)

        # Стилизация для drag & drop
        self._base_style = """
            QScrollArea {
                border: 2px dashed #cccccc;
                border-radius: 5px;
                background-color: #fafafa;
            }
            QScrollArea[dragOver="true"] {
                border: 2px dashed #4CAF50;
                background-color: #e8f5e8;
            }
        """
        self.setStyleSheet(self._base_style)

    def disable_hint_style(self):
        """Удаляет базовую рамку, оставляя стиль подсветки при перетаскивании."""
        self.setStyleSheet(
            """
            QScrollArea[dragOver="true"] {
                border: 2px dashed #4CAF50;
                background-color: #e8f5e8;
            }
        """
        )

    def dragEnterEvent(self, event):
        """Обработка входа перетаскиваемых файлов в область."""
        print("=== DRAG ENTER EVENT ===")
        print(f"Mime data has URLs: {event.mimeData().hasUrls()}")

        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            print(f"Number of URLs: {len(urls)}")

            all_paths = []
            xml_paths = []

            for url in urls:
                path = url.toLocalFile()
                all_paths.append(path)
                print(f"Path: '{path}'")

                # Более гибкая проверка XML файлов
                if path.lower().endswith(".xml") or path.lower().endswith(".XML"):
                    xml_paths.append(path)
                    print(f"  -> Valid XML file")
                else:
                    print(f"  -> Not an XML file")

            print(f"Total paths: {len(all_paths)}, XML paths: {len(xml_paths)}")

            if xml_paths:
                print("Accepting drag operation")
                event.acceptProposedAction()
                # Визуальная обратная связь
                self.setProperty("dragOver", True)
                self.style().unpolish(self)
                self.style().polish(self)
                return

        print("Ignoring drag operation")
        event.ignore()

    def dragMoveEvent(self, event):
        """Обработка движения перетаскиваемых файлов над областью."""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        """Обработка выхода перетаскиваемых файлов из области."""
        print("=== DRAG LEAVE EVENT ===")
        # Убираем визуальную обратную связь
        self.setProperty("dragOver", False)
        self.style().unpolish(self)
        self.style().polish(self)

    def dropEvent(self, event):
        """Обработка сброса файлов в область."""
        print("=== DROP EVENT ===")

        # Убираем визуальную обратную связь
        self.setProperty("dragOver", False)
        self.style().unpolish(self)
        self.style().polish(self)

        if not event.mimeData().hasUrls():
            print("No URLs in drop event")
            event.ignore()
            return

        urls = event.mimeData().urls()
        print(f"Processing {len(urls)} dropped URLs")

        all_paths = []
        xml_paths = []

        for i, url in enumerate(urls):
            path = url.toLocalFile()
            all_paths.append(path)
            print(f"URL {i + 1}: '{path}'")

            # Проверяем существование файла
            try:
                import os

                if not os.path.exists(path):
                    print(f"  -> File does not exist!")
                    continue

                if not os.path.isfile(path):
                    print(f"  -> Not a file!")
                    continue

                print(f"  -> File exists, size: {os.path.getsize(path)} bytes")
            except Exception as e:
                print(f"  -> Error checking file: {e}")
                continue

            # Более гибкая проверка XML
            if path.lower().endswith((".xml", ".XML")):
                xml_paths.append(path)
                print(f"  -> Added as XML file")
            else:
                # Пробуем проверить содержимое файла
                try:
                    with open(path, "r", encoding="utf-8", errors="ignore") as f:
                        first_line = f.readline().strip()
                        if first_line.startswith("<?xml") or "<" in first_line:
                            xml_paths.append(path)
                            print(f"  -> Added as XML file (detected by content)")
                        else:
                            print(f"  -> Not an XML file (content check)")
                except Exception as e:
                    print(f"  -> Could not check file content: {e}")

        print(f"Final result: {len(xml_paths)} XML files out of {len(all_paths)} total")

        if xml_paths:
            print("Calling callback with XML files...")
            try:
                self._callback(xml_paths)
                event.acceptProposedAction()
                print("Callback completed successfully")
            except Exception as e:
                print(f"Error in callback: {e}")
                QMessageBox.critical(
                    None, "Ошибка", f"Ошибка при обработке файлов: {e}"
                )
        else:
            print("No valid XML files found")
            if all_paths:
                QMessageBox.warning(
                    None,
                    "Предупреждение",
                    f"Среди {len(all_paths)} перетащенных файлов не найдено ни одного XML файла.\n"
                    "Поддерживаются только файлы с расширением .xml",
                )
            event.ignore()


class TranslationCostCalculator(QMainWindow):
    """Главное окно приложения"""

    def __init__(self):
        super().__init__()
        self.language_pairs: Dict[str, LanguagePairWidget] = {}
        self.pair_headers: Dict[str, str] = {}
        self.lang_display_ru: bool = True
        self._languages: List[Dict[str, str]] = load_languages()
        self.pm_managers, self.pm_last_index = load_pm_history()
        if 0 <= self.pm_last_index < len(self.pm_managers):
            self.current_pm = self.pm_managers[self.pm_last_index]
        else:
            self.current_pm = {"name_ru": "", "name_en": "", "email": ""}
        self.only_new_repeats_mode = False
        self.legal_entities = load_legal_entities()
        self.currency_symbol = CURRENCY_SYMBOLS.get("RUB", "₽")
        self.setup_ui()
        self.setup_style()
        self._preload_excel_app()

    def setup_ui(self):
        self.setGeometry(100, 100, 1000, 600)
        self.setMinimumSize(600, 400)
        self.resize(1000, 650)
        self.update_title()

        # меню с действиями загрузки/сохранения проекта
        project_menu = self.menuBar().addMenu("Проект")
        save_action = QAction("Сохранить проект", self)
        save_action.triggered.connect(self.save_project)
        project_menu.addAction(save_action)
        load_action = QAction("Загрузить проект", self)
        load_action.triggered.connect(self.load_project)
        project_menu.addAction(load_action)

        export_menu = self.menuBar().addMenu("Экспорт")
        save_excel_action = QAction("Сохранить Excel", self)
        save_excel_action.triggered.connect(self.save_excel)
        export_menu.addAction(save_excel_action)
        save_pdf_action = QAction("Сохранить PDF", self)
        save_pdf_action.triggered.connect(self.save_pdf)
        export_menu.addAction(save_pdf_action)

        pm_action = QAction("Проджект менеджер", self)
        pm_action.triggered.connect(self.show_pm_dialog)
        self.menuBar().addAction(pm_action)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QHBoxLayout()
        splitter = QSplitter(Qt.Horizontal)

        left_panel = self.create_left_panel()
        right_panel = self.create_right_panel()

        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)

        # чуть уже левая панель
        splitter.setSizes([600, 960])

        main_layout.addWidget(splitter)
        central_widget.setLayout(main_layout)

    def _preload_excel_app(self) -> None:
        if sys.platform != "win32":
            return

        def _warmup() -> None:
            try:
                app = get_excel_app()
                close_excel_app(app)
            except Exception:
                pass

        threading.Thread(target=_warmup, daemon=True).start()

    # ---------- LEFT ----------
    def create_left_panel(self) -> QWidget:
        container = QWidget()
        lay = QVBoxLayout()

        # Проект
        project_group = QGroupBox("Информация о проекте")
        p = QVBoxLayout()
        p.addWidget(QLabel("Название проекта:"))
        self.project_name_edit = QLineEdit()
        p.addWidget(self.project_name_edit)
        p.addWidget(QLabel("Название клиента:"))
        self.client_name_edit = QLineEdit()
        p.addWidget(self.client_name_edit)
        p.addWidget(QLabel("Контактное лицо:"))
        self.contact_person_edit = QLineEdit()
        p.addWidget(self.contact_person_edit)
        p.addWidget(QLabel("E-mail:"))
        self.email_edit = QLineEdit()
        p.addWidget(self.email_edit)
        p.addWidget(QLabel("Юрлицо:"))
        self.legal_entity_combo = QComboBox()
        self.legal_entity_combo.addItems(self.legal_entities.keys())
        self.legal_entity_combo.currentTextChanged.connect(self.on_legal_entity_changed)
        p.addWidget(self.legal_entity_combo)
        p.addWidget(QLabel("Валюта:"))
        self.currency_combo = QComboBox()
        self.currency_combo.addItems(["RUB", "EUR", "USD"])
        self.currency_combo.currentTextChanged.connect(self.on_currency_changed)
        p.addWidget(self.currency_combo)
        p.addWidget(QLabel("НДС, %:"))
        self.vat_spin = QDoubleSpinBox()
        self.vat_spin.setDecimals(2)
        self.vat_spin.setRange(0, 100)
        self.vat_spin.setValue(20.0)
        p.addWidget(self.vat_spin)
        project_group.setLayout(p)
        lay.addWidget(project_group)
        self.on_legal_entity_changed(self.legal_entity_combo.currentText())

        # Языковые пары
        self.pairs_group = QGroupBox(tr("Языковые пары", "ru"))
        pg = QVBoxLayout()

        # Переключатель RU/EN
        mode = QHBoxLayout()
        mode.addWidget(QLabel("Названия языков:"))
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

        # Добавление пары
        add_pair = QHBoxLayout()
        self.source_lang_combo = self._make_lang_combo()
        self.source_lang_combo.setEditable(True)
        add_pair.addWidget(self.source_lang_combo)
        add_pair.addWidget(QLabel("→"))
        self.target_lang_combo = self._make_lang_combo()
        self.target_lang_combo.setEditable(True)
        add_pair.addWidget(self.target_lang_combo)
        pg.addLayout(add_pair)

        self.add_pair_btn = QPushButton("Добавить языковую пару")
        self.add_pair_btn.clicked.connect(self.add_language_pair)
        pg.addWidget(self.add_pair_btn)

        pg.addWidget(QLabel("Текущие пары:"))
        self.pairs_list = QTextEdit()
        self.pairs_list.setMaximumHeight(110)
        self.pairs_list.setReadOnly(True)
        pg.addWidget(self.pairs_list)

        info_layout = QHBoxLayout()
        self.language_pairs_count_label = QLabel("Загружено языковых пар: 0")
        info_layout.addWidget(self.language_pairs_count_label)
        info_layout.addStretch()
        self.clear_pairs_btn = QPushButton("Очистить")
        self.clear_pairs_btn.clicked.connect(self.clear_language_pairs)
        info_layout.addWidget(self.clear_pairs_btn)
        pg.addLayout(info_layout)

        setup_layout = QHBoxLayout()
        self.project_setup_label = QLabel(
            tr("Запуск и управление проектом", "ru") + ":"
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

        # Добавление языка в справочник (без кода)
        add_lang_group = QGroupBox("Добавить язык в справочник")
        lg = QVBoxLayout()
        r1 = QHBoxLayout()
        r1.addWidget(QLabel("Название RU:"))
        self.new_lang_ru = QLineEdit()
        self.new_lang_ru.setPlaceholderText("Персидский")
        r1.addWidget(self.new_lang_ru)
        lg.addLayout(r1)
        r2 = QHBoxLayout()
        r2.addWidget(QLabel("Название EN:"))
        self.new_lang_en = QLineEdit()
        self.new_lang_en.setPlaceholderText("Persian")
        r2.addWidget(self.new_lang_en)
        lg.addLayout(r2)
        self.btn_add_lang = QPushButton("Добавить язык")
        self.btn_add_lang.clicked.connect(self.handle_add_language)
        lg.addWidget(self.btn_add_lang)
        add_lang_group.setLayout(lg)
        pg.addWidget(add_lang_group)

        self.pairs_group.setLayout(pg)
        lay.addWidget(self.pairs_group)

        lay.addStretch()
        container.setLayout(lay)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(container)
        scroll.setMinimumWidth(260)
        return scroll

    def _make_lang_combo(self) -> QComboBox:
        cb = QComboBox()
        self.populate_lang_combo(cb)
        return cb

    def populate_lang_combo(self, combo: QComboBox):
        """Заполняет комбобокс по _languages; при смене режима RU/EN старается сохранить выбор."""
        prev_text = combo.currentText() if combo.isEditable() else ""
        prev_idx = combo.currentIndex()
        prev_obj = combo.itemData(prev_idx) if prev_idx >= 0 else None

        combo.blockSignals(True)
        combo.clear()
        for lang in self._languages:
            name = lang["ru"] if self.lang_display_ru else lang["en"]
            name = shorten_locale(name, "ru" if self.lang_display_ru else "en")
            label = f"{name}"
            combo.addItem(label, lang)
        combo.blockSignals(False)

        # восстановление выбора по объекту (en/ru)
        if isinstance(prev_obj, dict):
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

    def on_lang_mode_changed(self, value: int):
        self.lang_display_ru = value == 1
        self.populate_lang_combo(self.source_lang_combo)
        self.populate_lang_combo(self.target_lang_combo)
        lang = "ru" if self.lang_display_ru else "en"
        if getattr(self, "project_setup_widget", None):
            self.project_setup_widget.set_language(lang)
        if getattr(self, "additional_services_widget", None):
            self.additional_services_widget.set_language(lang)
        self.project_setup_label.setText(tr("Запуск и управление проектом", lang) + ":")
        self.pairs_group.setTitle(tr("Языковые пары", lang))
        for pair_key, widget in self.language_pairs.items():
            widget.set_language(lang)
            display_name = self._display_pair_name(pair_key)
            widget.set_pair_name(display_name)
            right_key = pair_key.split(" → ")[1]
            lang_info = self._find_language_by_key(right_key)
            self.pair_headers[pair_key] = lang_info[lang]
        self.update_pairs_list()
        self.tabs.setTabText(0, tr("Языковые пары", lang))
        self.tabs.setTabText(1, tr("Дополнительные услуги", lang))

    def on_currency_changed(self, code: str):
        self.currency_symbol = CURRENCY_SYMBOLS.get(code, code)
        if getattr(self, "project_setup_widget", None):
            self.project_setup_widget.set_currency(self.currency_symbol, code)
        for w in self.language_pairs.values():
            w.set_currency(self.currency_symbol, code)
        if getattr(self, "additional_services_widget", None):
            self.additional_services_widget.set_currency(self.currency_symbol, code)

    def on_legal_entity_changed(self, entity: str):
        is_art = entity == "Арт"
        self.vat_spin.setEnabled(is_art)
        if is_art and self.vat_spin.value() == 0:
            self.vat_spin.setValue(20.0)
        if not is_art:
            self.vat_spin.setValue(0.0)

    # ---------- RIGHT ----------
    def create_right_panel(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout()
        self.tabs = QTabWidget()

        # Создаем основной скроллируемый контейнер для языковых пар
        self.pairs_scroll = QScrollArea()
        self.pairs_scroll.setWidgetResizable(True)
        self.pairs_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.pairs_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        # Создаем виджет-контейнер для языковых пар
        self.pairs_container_widget = QWidget()
        self.pairs_layout = QVBoxLayout()

        self.only_new_repeats_btn = QPushButton("Только новые слова и повторы")
        self.only_new_repeats_btn.clicked.connect(self.toggle_only_new_repeats_mode)
        self.pairs_layout.addWidget(self.only_new_repeats_btn)

        # Таблица запуска и управления проектом
        self.project_setup_widget = ProjectSetupWidget(
            self.project_setup_fee_spin.value(),
            self.currency_symbol,
            self.currency_combo.currentText(),
            lang="ru" if self.lang_display_ru else "en",
        )
        self.project_setup_widget.remove_requested.connect(
            self.remove_project_setup_widget
        )
        self.pairs_layout.addWidget(self.project_setup_widget)
        self.project_setup_fee_spin.valueChanged.connect(
            self.update_project_setup_volume_from_spin
        )
        self.project_setup_widget.table.itemChanged.connect(
            self.on_project_setup_item_changed
        )

        # Добавляем подсказку для пользователя
        self.drop_hint_label = QLabel(
            "Перетащите XML файлы отчетов Trados сюда для автоматического заполнения"
        )
        self.drop_hint_label.setStyleSheet(
            """
            QLabel {
                color: #666666;
                font-style: italic;
                padding: 20px;
                text-align: center;
            }
        """
        )
        self.drop_hint_label.setAlignment(Qt.AlignCenter)
        self.pairs_layout.addWidget(self.drop_hint_label)

        # Добавляем растягивающийся элемент в конце
        self.pairs_layout.addStretch()

        self.pairs_container_widget.setLayout(self.pairs_layout)
        self.pairs_scroll.setWidget(self.pairs_container_widget)

        # Настраиваем drag & drop для скроллируемой области
        self.pairs_scroll.setAcceptDrops(True)
        self.setup_drag_drop()

        self.tabs.addTab(self.pairs_scroll, tr("Языковые пары", "ru"))

        self.additional_services_widget = AdditionalServicesWidget(
            self.currency_symbol,
            self.currency_combo.currentText(),
            lang="ru" if self.lang_display_ru else "en",
        )
        add_scroll = QScrollArea()
        add_scroll.setWidget(self.additional_services_widget)
        add_scroll.setWidgetResizable(True)
        self.tabs.addTab(add_scroll, tr("Дополнительные услуги", "ru"))

        lay.addWidget(self.tabs)
        w.setLayout(lay)
        return w

    def setup_drag_drop(self):
        drop_area = DropArea(self.handle_xml_drop)

        # Переносим содержимое в новую область
        drop_area.setWidget(self.pairs_container_widget)

        # Обновляем вкладку
        self.tabs.removeTab(0)
        self.tabs.insertTab(0, drop_area, tr("Языковые пары", "ru"))
        self.pairs_scroll = drop_area

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
        if self.only_new_repeats_mode:
            self.only_new_repeats_btn.setText("Показать 4 строки")
        else:
            self.only_new_repeats_btn.setText("Только новые слова и повторы")

    def setup_style(self):
        self.setStyleSheet(APP_STYLE)

    def update_title(self):
        name = self.current_pm.get("name_ru") or self.current_pm.get("name_en") or ""
        if name:
            self.setWindowTitle(f"RateApp - {name}")
        else:
            self.setWindowTitle("RateApp")

    def show_pm_dialog(self):
        dlg = ProjectManagerDialog(self.pm_managers, self.pm_last_index, self)
        if dlg.exec():
            self.pm_managers, self.pm_last_index = dlg.result()
            save_pm_history(self.pm_managers, self.pm_last_index)
            if 0 <= self.pm_last_index < len(self.pm_managers):
                self.current_pm = self.pm_managers[self.pm_last_index]
            else:
                self.current_pm = {"name_ru": "", "name_en": "", "email": ""}
            self.update_title()

    # ---------- PROJECT SETUP ----------
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

    # ---------- LANG ADD ----------
    def handle_add_language(self):
        ru = (self.new_lang_ru.text() or "").strip()
        en = (self.new_lang_en.text() or "").strip()

        if not (ru or en):
            QMessageBox.warning(
                self, "Ошибка", "Укажите хотя бы одно название (RU или EN)."
            )
            return

        if add_language(en, ru):
            QMessageBox.information(self, "Готово", "Язык сохранён в конфиг.")
            self._languages = load_languages()
            self.populate_lang_combo(self.source_lang_combo)
            self.populate_lang_combo(self.target_lang_combo)
            # очистим поля
            self.new_lang_ru.clear()
            self.new_lang_en.clear()
        else:
            QMessageBox.critical(self, "Ошибка", "Не удалось сохранить язык в конфиг.")

    # ---------- ACTIONS ----------
    def _parse_combo(self, combo: QComboBox) -> Dict[str, Any]:
        """Если текущий текст совпадает с одним из элементов — это словарный; иначе кастом."""
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

    def add_language_pair(self):
        src = self._parse_combo(self.source_lang_combo)
        tgt = self._parse_combo(self.target_lang_combo)
        if not src["text"] or not tgt["text"]:
            QMessageBox.warning(self, "Ошибка", "Выберите/введите оба языка")
            return

        # внутренний ключ пары фиксируем по EN (если есть), иначе по RU/тексту — стабильно при смене режима
        def key_name(obj: Dict[str, Any]) -> str:
            return obj["en"] or obj["ru"] or obj["text"]

        left_key = key_name(src)
        right_key = key_name(tgt)
        pair_key = f"{left_key} → {right_key}"

        # Видимое название пары берём непосредственно из полей ввода,
        # чтобы оно соответствовало выбору пользователя (RU/EN)
        display_name = f"{src['text']} - {tgt['text']}"

        if pair_key in self.language_pairs:
            QMessageBox.warning(self, "Ошибка", "Такая языковая пара уже существует")
            return

        # Заголовок для Excel = текущее отображаемое имя целевого
        if tgt["dict"]:
            name = tgt["ru"] if self.lang_display_ru else tgt["en"]
            header_title = shorten_locale(
                name, "ru" if self.lang_display_ru else "en"
            )
        else:
            header_title = tgt["text"]
        self.pair_headers[pair_key] = header_title

        widget = LanguagePairWidget(
            display_name,
            self.currency_symbol,
            self.currency_combo.currentText(),
            lang="ru" if self.lang_display_ru else "en",
        )  # только Перевод
        widget.remove_requested.connect(
            lambda pk=pair_key: self.remove_language_pair(pk)
        )
        self.language_pairs[pair_key] = widget
        if self.only_new_repeats_mode:
            widget.set_only_new_and_repeats_mode(True)

        # Вставляем новый виджет перед растягивающимся элементом
        self.pairs_layout.insertWidget(self.pairs_layout.count() - 1, widget)

        self.update_pairs_list()

        self.source_lang_combo.setCurrentIndex(0)
        self.target_lang_combo.setCurrentIndex(0)

    def _pair_sort_key(self, pair_key: str) -> str:
        parts = pair_key.split(" → ")
        return parts[1] if len(parts) > 1 else parts[0]

    def _find_language_by_key(self, key: str) -> Dict[str, str]:
        for lang in self._languages:
            if key.lower() == lang["en"].lower() or key.lower() == lang["ru"].lower():
                return lang

        try:
            tag = key
            match = re.match(r"(.+?)\s*\((.+)\)$", key)
            if match:
                base = match.group(1).strip()
                region = match.group(2).strip()
                base_code = langcodes.find(base)
                try:
                    region_code = pycountry.countries.lookup(region).alpha_2
                except LookupError:
                    region_code = region
                tag = f"{base_code}-{region_code}".lower()
            elif re.fullmatch(r"[A-Za-z]{2,3}(?:-[A-Za-z]{2,3})?", key):
                tag = key.lower()
            else:
                tag = langcodes.find(key)
            lang_obj = langcodes.Language.get(tag)
            en_name = lang_obj.display_name("en").title()
            en_name = shorten_locale(en_name, "en")
            ru_name = lang_obj.display_name("ru")
            ru_name = ru_name[0].upper() + ru_name[1:]
            ru_name = shorten_locale(ru_name, "ru")
            return {"en": en_name, "ru": ru_name}
        except Exception:
            return {"en": key, "ru": key}

    def _display_pair_name(self, pair_key: str) -> str:
        left_key, right_key = pair_key.split(" → ")
        left = self._find_language_by_key(left_key)
        right = self._find_language_by_key(right_key)
        lang = "ru" if self.lang_display_ru else "en"
        return f"{left[lang]} - {right[lang]}"

    def update_pairs_list(self):
        self.pairs_list.setText(
            "\n".join(
                f"{w.pair_name}   [заголовок: {self.pair_headers.get(key, w.pair_name)}]"
                for key, w in sorted(
                    self.language_pairs.items(),
                    key=lambda kv: self._pair_sort_key(kv[0]),
                )
            )
        )
        pair_count = len(self.language_pairs)
        self.language_pairs_count_label.setText(f"Загружено языковых пар: {pair_count}")
        auto_fee = max(0.5, round((pair_count + 1) / 4 * 4) / 4)
        self.project_setup_fee_spin.blockSignals(True)
        self.project_setup_fee_spin.setValue(auto_fee)
        self.project_setup_fee_spin.blockSignals(False)
        self.update_project_setup_volume_from_spin(auto_fee)

    def remove_language_pair(self, pair_key: str):
        widget = self.language_pairs.pop(pair_key, None)
        if widget:
            widget.setParent(None)
            self.pair_headers.pop(pair_key, None)
        self.update_pairs_list()

    def clear_language_pairs(self):
        for w in self.language_pairs.values():
            w.setParent(None)
        self.language_pairs.clear()
        self.pair_headers.clear()
        self.update_pairs_list()

    def handle_xml_drop(self, paths: List[str], replace: bool = False):
        """Улучшенная обработка XML файлов с детальным логированием."""
        print(f"\n=== HANDLING XML DROP ===")
        print(f"Received {len(paths)} XML files")
        print(f"Replace mode: {replace}")

        try:
            data, warnings = parse_reports(paths)
            print(f"Parse results: {len(data)} language pairs found")

            if warnings:
                warning_msg = f"Предупреждения при обработке файлов:\n" + "\n".join(
                    warnings
                )
                print(f"Warnings: {warning_msg}")
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
                print(f"\nProcessing pair: {pair_key}")
                print(f"Volumes: {volumes}")

                widget = self.language_pairs.get(pair_key)

                # Отображаемое имя и заголовок формируем в соответствии с текущим режимом
                display_name = self._display_pair_name(pair_key)
                tgt_key = pair_key.split(" → ")[1] if " → " in pair_key else pair_key
                lang_info = self._find_language_by_key(tgt_key)
                header_title = (
                    lang_info["ru"] if self.lang_display_ru else lang_info["en"]
                )

                if widget is None:
                    print(f"Creating new widget for pair: {pair_key}")
                    widget = LanguagePairWidget(
                        display_name,
                        self.currency_symbol,
                        self.currency_combo.currentText(),
                        lang="ru" if self.lang_display_ru else "en",
                    )
                    widget.remove_requested.connect(
                        lambda pk=pair_key: self.remove_language_pair(pk)
                    )
                    self.language_pairs[pair_key] = widget
                    # Вставляем новый виджет перед растягивающимся элементом
                    self.pairs_layout.insertWidget(
                        self.pairs_layout.count() - 1, widget
                    )
                    self.pair_headers[pair_key] = header_title
                    added_pairs += 1
                else:
                    print(f"Updating existing widget for pair: {pair_key}")
                    widget.set_pair_name(display_name)
                    widget.set_language("ru" if self.lang_display_ru else "en")
                    self.pair_headers[pair_key] = header_title
                    updated_pairs += 1

                if self.only_new_repeats_mode:
                    widget.set_only_new_and_repeats_mode(True)

                # Обновляем данные в таблице перевода
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
                        print(f"  Set new words: {new_total}")
                        print(f"  Set repeats: {repeat_total}")
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
                        print(
                            f"  Updated new words: {prev_new} + {new_total} = {prev_new + new_total}"
                        )
                        print(
                            f"  Updated repeats: {prev_rep} + {repeat_total} = {prev_rep + repeat_total}"
                        )
                else:
                    total_volume = 0
                    for idx, row_info in enumerate(ServiceConfig.TRANSLATION_ROWS):
                        row_name = row_info["name"]
                        add_val = volumes.get(row_name, 0)
                        total_volume += add_val

                        if replace:
                            table.item(idx, 1).setText(str(add_val))
                            print(f"  Set {row_name}: {add_val}")
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
                                print(
                                    f"  Updated {row_name}: {prev} + {add_val} = {new_val}"
                                )
                            except (ValueError, TypeError) as e:
                                print(f"  Error updating {row_name}: {e}")
                                table.item(idx, 1).setText(str(add_val))

                print(f"  Total volume for {pair_key}: {total_volume}")

                # Пересчитываем ставки и суммы
                widget.update_rates_and_sums(
                    table, group.rows_config, group.base_rate_row
                )

            # Обновляем список пар
            # ensure widgets are ordered alphabetically by target language
            sorted_items = sorted(
                self.language_pairs.items(), key=lambda kv: self._pair_sort_key(kv[0])
            )
            # remove existing widgets and re-insert in sorted order
            for w in self.language_pairs.values():
                self.pairs_layout.removeWidget(w)
            for _, w in sorted_items:
                self.pairs_layout.insertWidget(self.pairs_layout.count() - 1, w)
            self.language_pairs = dict(sorted_items)

            self.update_pairs_list()

            # Показываем результат пользователю
            result_msg = f"Обработка завершена!\n\n"
            if added_pairs > 0:
                result_msg += f"Добавлено новых языковых пар: {added_pairs}\n"
            if updated_pairs > 0:
                result_msg += f"Обновлено существующих пар: {updated_pairs}\n"
            result_msg += f"\nВсего обработано языковых пар: {len(data)}"

        except Exception as e:
            error_msg = f"Ошибка при обработке XML файлов: {str(e)}"
            print(f"ERROR: {error_msg}")
            QMessageBox.critical(self, "Ошибка", error_msg)

    def collect_project_data(self) -> Dict[str, Any]:
        data = {
            "project_name": self.project_name_edit.text(),
            "client_name": self.client_name_edit.text(),
            "contact_person": self.contact_person_edit.text(),
            "email": self.email_edit.text(),
            "legal_entity": self.legal_entity_combo.currentText(),
            "currency": self.currency_combo.currentText(),
            "language_pairs": [],
            "additional_services": [],
            "pm_name": self.current_pm.get(
                "name_ru" if self.lang_display_ru else "name_en", ""
            ),
            "pm_email": self.current_pm.get("email", ""),
            "project_setup_fee": self.project_setup_fee_spin.value(),
            "project_setup": (
                self.project_setup_widget.get_data()
                if self.project_setup_widget
                else []
            ),
            "vat_rate": self.vat_spin.value() if self.vat_spin.isEnabled() else 0,
        }
        for pair_key, pair_widget in self.language_pairs.items():
            p = pair_widget.get_data()
            if p["services"]:
                p["header_title"] = self.pair_headers.get(
                    pair_key, pair_widget.pair_name
                )
                data["language_pairs"].append(p)
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
        return data

    def save_excel(self):
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

        client_name = project_data["client_name"].replace(" ", "_")
        entity_for_file = self.legal_entity_combo.currentText().replace(" ", "_")
        currency = self.currency_combo.currentText()
        date_str = datetime.now().strftime("%Y-%m-%d")
        filename = f"{date_str}-{entity_for_file}-{currency}-{client_name}.xlsx"

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить Excel файл", filename, "Excel files (*.xlsx)"
        )
        if not file_path:
            return

        entity_name = self.legal_entity_combo.currentText()
        template_path = self.legal_entities.get(entity_name)
        exporter = ExcelExporter(
            template_path,
            currency=self.currency_combo.currentText(),
            lang="ru" if self.lang_display_ru else "en",
        )
        with Progress(parent=self) as progress:
            success = exporter.export_to_excel(
                project_data, file_path, progress_callback=progress.on_progress
            )
        if success:
            QMessageBox.information(self, "Успех", f"Файл сохранен: {file_path}")
        else:
            QMessageBox.critical(self, "Ошибка", "Не удалось сохранить файл")

    def save_pdf(self):
        if not self.project_name_edit.text().strip():
            QMessageBox.warning(self, "Ошибка", "Введите название проекта")
            return
        if not self.client_name_edit.text().strip():
            QMessageBox.warning(self, "Ошибка", "Введите название клиента")
            return
        project_data = self.collect_project_data()
        client_name = project_data["client_name"].replace(" ", "_")
        entity_for_file = self.legal_entity_combo.currentText().replace(" ", "_")
        currency = self.currency_combo.currentText()
        date_str = datetime.now().strftime("%Y-%m-%d")
        filename = f"{date_str}-{entity_for_file}-{currency}-{client_name}.pdf"
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить PDF файл", filename, "PDF files (*.pdf)"
        )
        if not file_path:
            return
        template_path = self.legal_entities.get(self.legal_entity_combo.currentText())
        exporter = ExcelExporter(
            template_path,
            currency=currency,
            lang="ru" if self.lang_display_ru else "en",
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
                    if not xlsx_to_pdf(xlsx_path, pdf_path):
                        QMessageBox.critical(
                            self, "Ошибка", "Не удалось конвертировать в PDF"
                        )
                        return
                    progress.set_value(100)
                    shutil.copyfile(pdf_path, file_path)
                QMessageBox.information(self, "Успех", f"Файл сохранен: {file_path}")
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
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(project_data, f, ensure_ascii=False, indent=2)
            QMessageBox.information(self, "Успех", f"Проект сохранен: {file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить проект: {e}")

    def load_project(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Загрузить проект", "", "JSON files (*.json)"
        )
        if not file_path:
            return
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                project_data = json.load(f)
            self.load_project_data(project_data)
            QMessageBox.information(self, "Успех", "Проект загружен")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить проект: {e}")

    def load_project_data(self, project_data: Dict[str, Any]):
        self.project_name_edit.setText(project_data.get("project_name", ""))
        self.client_name_edit.setText(project_data.get("client_name", ""))
        self.contact_person_edit.setText(project_data.get("contact_person", ""))
        self.email_edit.setText(project_data.get("email", ""))
        if hasattr(self, "legal_entity_combo"):
            self.legal_entity_combo.setCurrentText(project_data.get("legal_entity", ""))
        if hasattr(self, "currency_combo"):
            self.currency_combo.setCurrentText(project_data.get("currency", "RUB"))
        if hasattr(self, "vat_spin"):
            self.vat_spin.setValue(project_data.get("vat_rate", 20.0))
            self.on_legal_entity_changed(self.legal_entity_combo.currentText())

        for w in self.language_pairs.values():
            w.setParent(None)
        self.language_pairs.clear()
        self.pair_headers.clear()

        for pair_data in project_data.get("language_pairs", []):
            pair_key = pair_data[
                "pair_name"
            ]  # в твоём текущем формате это строка, оставляем как есть
            header_title = pair_data.get("header_title", pair_key)
            widget = LanguagePairWidget(
                pair_key, self.currency_symbol, self.currency_combo.currentText()
            )
            widget.remove_requested.connect(
                lambda pk=pair_key: self.remove_language_pair(pk)
            )
            self.language_pairs[pair_key] = widget
            # Вставляем новый виджет перед растягивающимся элементом
            self.pairs_layout.insertWidget(self.pairs_layout.count() - 1, widget)
            self.pair_headers[pair_key] = header_title

            services = pair_data.get("services", {})
            if "translation" in services:
                widget.translation_group.setChecked(True)
                widget.load_table_data(services["translation"])

        self.update_pairs_list()

        ps_rows = project_data.get("project_setup")
        if ps_rows:
            if not self.project_setup_widget:
                self.project_setup_widget = ProjectSetupWidget(
                    self.project_setup_fee_spin.value(),
                    self.currency_symbol,
                    self.currency_combo.currentText(),
                )
                self.project_setup_widget.remove_requested.connect(
                    self.remove_project_setup_widget
                )
                self.project_setup_widget.table.itemChanged.connect(
                    self.on_project_setup_item_changed
                )
                self.pairs_layout.insertWidget(0, self.project_setup_widget)
            self.project_setup_widget.load_data(ps_rows)
            first_vol = ps_rows[0].get("volume", 0)
            self.project_setup_fee_spin.blockSignals(True)
            self.project_setup_fee_spin.setValue(first_vol)
            self.project_setup_fee_spin.blockSignals(False)
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
