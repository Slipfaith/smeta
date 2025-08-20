import json
from datetime import datetime
from typing import Dict, List, Any

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QGroupBox, QTextEdit, QFileDialog, QMessageBox, QScrollArea, QTabWidget, QSplitter,
    QFrame, QComboBox, QSlider
)
from PySide6.QtCore import Qt

from gui.language_pair import LanguagePairWidget
from gui.additional_services import AdditionalServicesWidget
from gui.styles import APP_STYLE
from logic.excel_exporter import ExcelExporter
from logic.user_config import load_languages, add_language
from logic.trados_xml_parser import parse_reports
from logic.service_config import ServiceConfig


class DropWidget(QWidget):
    """Принимает перетаскивание XML-файлов и отдаёт список путей в колбек."""

    def __init__(self, callback, parent=None):
        super().__init__(parent)
        self._callback = callback
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):  # noqa: D401
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.toLocalFile().lower().endswith('.xml'):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event):  # noqa: D401
        paths = [url.toLocalFile() for url in event.mimeData().urls() if url.toLocalFile().lower().endswith('.xml')]
        if paths:
            self._callback(paths)
            event.acceptProposedAction()
        else:
            event.ignore()

class TranslationCostCalculator(QMainWindow):
    """Главное окно приложения"""

    def __init__(self):
        super().__init__()
        self.language_pairs: Dict[str, LanguagePairWidget] = {}
        self.pair_headers: Dict[str, str] = {}  # pair_key -> header_title (RU/EN целевого или кастом)
        self.lang_display_ru: bool = True       # True=RU, False=EN
        self._languages: List[Dict[str, str]] = load_languages()
        self.setup_ui()
        self.setup_style()

    def setup_ui(self):
        self.setWindowTitle("Калькулятор стоимости переводческих проектов")
        self.setGeometry(100, 100, 1000, 600)
        self.setMinimumSize(900, 600)
        self.resize(1000, 650)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QHBoxLayout()
        splitter = QSplitter(Qt.Horizontal)

        left_panel = self.create_left_panel()
        right_panel = self.create_right_panel()

        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)

        # чуть уже левая панель
        splitter.setSizes([240, 960])

        main_layout.addWidget(splitter)
        central_widget.setLayout(main_layout)

    # ---------- LEFT ----------
    def create_left_panel(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout()

        # Проект
        project_group = QGroupBox("Информация о проекте")
        p = QVBoxLayout()
        p.addWidget(QLabel("Название проекта:"))
        self.project_name_edit = QLineEdit(); p.addWidget(self.project_name_edit)
        p.addWidget(QLabel("Номер/код проекта:"))
        self.project_code_edit = QLineEdit(); p.addWidget(self.project_code_edit)
        p.addWidget(QLabel("Название клиента:"))
        self.client_name_edit = QLineEdit(); p.addWidget(self.client_name_edit)
        p.addWidget(QLabel("Контактное лицо:"))
        self.contact_person_edit = QLineEdit(); p.addWidget(self.contact_person_edit)
        p.addWidget(QLabel("E-mail:"))
        self.email_edit = QLineEdit(); p.addWidget(self.email_edit)
        project_group.setLayout(p); lay.addWidget(project_group)

        # Языковые пары
        pairs_group = QGroupBox("Языковые пары")
        pg = QVBoxLayout()

        # Переключатель RU/EN
        mode = QHBoxLayout()
        mode.addWidget(QLabel("Названия языков:")); mode.addStretch(1)
        mode.addWidget(QLabel("EN"))
        self.lang_mode_slider = QSlider(Qt.Horizontal); self.lang_mode_slider.setRange(0, 1); self.lang_mode_slider.setValue(1)
        self.lang_mode_slider.setFixedWidth(70); self.lang_mode_slider.valueChanged.connect(self.on_lang_mode_changed)
        mode.addWidget(self.lang_mode_slider); mode.addWidget(QLabel("RU"))
        pg.addLayout(mode)

        # Добавление пары
        add_pair = QHBoxLayout()
        self.source_lang_combo = self._make_lang_combo(); self.source_lang_combo.setEditable(True); add_pair.addWidget(self.source_lang_combo)
        add_pair.addWidget(QLabel("→"))
        self.target_lang_combo = self._make_lang_combo(); self.target_lang_combo.setEditable(True); add_pair.addWidget(self.target_lang_combo)
        pg.addLayout(add_pair)

        self.add_pair_btn = QPushButton("Добавить языковую пару")
        self.add_pair_btn.clicked.connect(self.add_language_pair)
        pg.addWidget(self.add_pair_btn)

        pg.addWidget(QLabel("Текущие пары:"))
        self.pairs_list = QTextEdit(); self.pairs_list.setMaximumHeight(110); self.pairs_list.setReadOnly(True)
        pg.addWidget(self.pairs_list)

        # Добавление языка в справочник (без кода)
        add_lang_group = QGroupBox("Добавить язык в справочник")
        lg = QVBoxLayout()
        r1 = QHBoxLayout(); r1.addWidget(QLabel("Название RU:")); self.new_lang_ru = QLineEdit(); self.new_lang_ru.setPlaceholderText("Персидский"); r1.addWidget(self.new_lang_ru); lg.addLayout(r1)
        r2 = QHBoxLayout(); r2.addWidget(QLabel("Название EN:")); self.new_lang_en = QLineEdit(); self.new_lang_en.setPlaceholderText("Persian");  r2.addWidget(self.new_lang_en); lg.addLayout(r2)
        self.btn_add_lang = QPushButton("Добавить язык"); self.btn_add_lang.clicked.connect(self.handle_add_language); lg.addWidget(self.btn_add_lang)
        add_lang_group.setLayout(lg); pg.addWidget(add_lang_group)

        pairs_group.setLayout(pg); lay.addWidget(pairs_group)

        # Действия
        actions = QGroupBox("Действия"); a = QVBoxLayout()
        t = QHBoxLayout(); self.template_path_edit = QLineEdit(); self.template_path_edit.setPlaceholderText("Путь к шаблону Excel"); t.addWidget(self.template_path_edit)
        btn = QPushButton("Выбрать"); btn.clicked.connect(self.select_template); t.addWidget(btn); a.addLayout(t)
        self.save_excel_btn = QPushButton("Сохранить Excel"); self.save_excel_btn.clicked.connect(self.save_excel); a.addWidget(self.save_excel_btn)
        self.save_pdf_btn = QPushButton("Сохранить PDF");   self.save_pdf_btn.clicked.connect(self.save_pdf);   a.addWidget(self.save_pdf_btn)
        a.addWidget(QFrame())
        self.save_project_btn = QPushButton("Сохранить проект"); self.save_project_btn.clicked.connect(self.save_project); a.addWidget(self.save_project_btn)
        self.load_project_btn = QPushButton("Загрузить проект"); self.load_project_btn.clicked.connect(self.load_project); a.addWidget(self.load_project_btn)
        actions.setLayout(a); lay.addWidget(actions)

        lay.addStretch()
        w.setLayout(lay)
        return w

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
            label = f"{name}"
            combo.addItem(label, lang)
        combo.blockSignals(False)

        # восстановление выбора по объекту (en/ru)
        if isinstance(prev_obj, dict):
            for i in range(combo.count()):
                d = combo.itemData(i)
                if isinstance(d, dict) and d.get("en") == prev_obj.get("en") and d.get("ru") == prev_obj.get("ru"):
                    combo.setCurrentIndex(i)
                    break
        elif prev_text:
            combo.setEditable(True)
            combo.setCurrentIndex(-1)
            combo.setEditText(prev_text)

    def on_lang_mode_changed(self, value: int):
        self.lang_display_ru = (value == 1)
        self.populate_lang_combo(self.source_lang_combo)
        self.populate_lang_combo(self.target_lang_combo)

    # ---------- RIGHT ----------
    def create_right_panel(self) -> QWidget:
        w = QWidget(); lay = QVBoxLayout()
        self.tabs = QTabWidget()

        self.pairs_scroll = QScrollArea()
        self.pairs_widget = DropWidget(self.handle_xml_drop)
        self.pairs_layout = QVBoxLayout()
        self.pairs_widget.setLayout(self.pairs_layout)
        self.pairs_scroll.setWidget(self.pairs_widget)
        self.pairs_scroll.setWidgetResizable(True)
        self.tabs.addTab(self.pairs_scroll, "Языковые пары")

        self.additional_services_widget = AdditionalServicesWidget()
        add_scroll = QScrollArea()
        add_scroll.setWidget(self.additional_services_widget)
        add_scroll.setWidgetResizable(True)
        self.tabs.addTab(add_scroll, "Дополнительные услуги")

        lay.addWidget(self.tabs); w.setLayout(lay)
        return w

    def setup_style(self):
        self.setStyleSheet(APP_STYLE)

    # ---------- LANG ADD ----------
    def handle_add_language(self):
        ru = (self.new_lang_ru.text() or "").strip()
        en = (self.new_lang_en.text() or "").strip()

        if not (ru or en):
            QMessageBox.warning(self, "Ошибка", "Укажите хотя бы одно название (RU или EN).")
            return

        if add_language(en, ru):
            QMessageBox.information(self, "Готово", "Язык сохранён в конфиг.")
            self._languages = load_languages()
            self.populate_lang_combo(self.source_lang_combo)
            self.populate_lang_combo(self.target_lang_combo)
            # очистим поля
            self.new_lang_ru.clear(); self.new_lang_en.clear()
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
                return {"en": data.get("en", ""), "ru": data.get("ru", ""), "text": text, "dict": True}
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
            header_title = (tgt["ru"] if self.lang_display_ru else tgt["en"])
        else:
            header_title = tgt["text"]
        self.pair_headers[pair_key] = header_title

        widget = LanguagePairWidget(display_name)  # только Перевод
        self.language_pairs[pair_key] = widget
        self.pairs_layout.addWidget(widget)

        self.update_pairs_list()

        self.source_lang_combo.setCurrentIndex(0)
        self.target_lang_combo.setCurrentIndex(0)

    def update_pairs_list(self):
        self.pairs_list.setText("\n".join(
            f"{w.pair_name}   [заголовок: {self.pair_headers.get(key, w.pair_name)}]"
            for key, w in self.language_pairs.items()
        ))

    def handle_xml_drop(self, paths: List[str], replace: bool = False):
        data, warnings = parse_reports(paths)
        if warnings:
            QMessageBox.warning(self, "Предупреждение", "\n".join(warnings))
        for pair_key, volumes in data.items():
            widget = self.language_pairs.get(pair_key)
            display_name = pair_key.replace(" → ", " - ")
            if widget is None:
                widget = LanguagePairWidget(display_name)
                self.language_pairs[pair_key] = widget
                self.pairs_layout.addWidget(widget)
                tgt = pair_key.split(" → ")[1] if " → " in pair_key else pair_key
                self.pair_headers[pair_key] = tgt
            group = widget.translation_group
            table = group.table
            group.setChecked(True)
            for idx, row_info in enumerate(ServiceConfig.TRANSLATION_ROWS):
                row_name = row_info["name"]
                add_val = volumes.get(row_name, 0)
                if replace:
                    table.item(idx, 1).setText(str(add_val))
                else:
                    prev = float((table.item(idx, 1).text() if table.item(idx, 1) else "0") or "0")
                    table.item(idx, 1).setText(str(prev + add_val))
            widget.update_rates_and_sums(table, group.rows_config, group.base_rate_row)
        self.update_pairs_list()

    def select_template(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Выбрать шаблон Excel", "", "Excel files (*.xlsx *.xls)")
        if file_path:
            self.template_path_edit.setText(file_path)

    def collect_project_data(self) -> Dict[str, Any]:
        data = {
            "project_name": self.project_name_edit.text(),
            "project_code": self.project_code_edit.text(),
            "client_name": self.client_name_edit.text(),
            "contact_person": self.contact_person_edit.text(),
            "email": self.email_edit.text(),
            "language_pairs": [],
            "additional_services": {}
        }
        for pair_key, pair_widget in self.language_pairs.items():
            p = pair_widget.get_data()
            if p["services"]:
                p["header_title"] = self.pair_headers.get(pair_key, pair_widget.pair_name)
                data["language_pairs"].append(p)
        additional = self.additional_services_widget.get_data()
        if additional:
            data["additional_services"] = additional
        return data

    def save_excel(self):
        if not self.client_name_edit.text().strip():
            QMessageBox.warning(self, "Ошибка", "Введите название клиента")
            return
        project_data = self.collect_project_data()
        if not project_data["language_pairs"] and not project_data["additional_services"]:
            QMessageBox.warning(self, "Ошибка", "Добавьте хотя бы одну услугу")
            return

        client_name = project_data["client_name"].replace(" ", "_")
        date_str = datetime.now().strftime("%Y%m%d")
        filename = f"КП_{client_name}_{date_str}.xlsx"

        file_path, _ = QFileDialog.getSaveFileName(self, "Сохранить Excel файл", filename, "Excel files (*.xlsx)")
        if not file_path:
            return

        template_path = self.template_path_edit.text().strip()
        exporter = ExcelExporter(template_path if template_path else None)

        if exporter.export_to_excel(project_data, file_path):
            QMessageBox.information(self, "Успех", f"Файл сохранен: {file_path}")
        else:
            QMessageBox.critical(self, "Ошибка", "Не удалось сохранить файл")

    def save_pdf(self):
        QMessageBox.information(
            self, "Функция PDF",
            "Функция сохранения в PDF будет реализована в следующей версии.\n"
            "Пока вы можете сохранить Excel файл и конвертировать его в PDF вручную."
        )

    def save_project(self):
        if not self.project_name_edit.text().strip():
            QMessageBox.warning(self, "Ошибка", "Введите название проекта")
            return
        project_data = self.collect_project_data()
        project_name = project_data["project_name"].replace(" ", "_")
        filename = f"Проект_{project_name}.json"
        file_path, _ = QFileDialog.getSaveFileName(self, "Сохранить проект", filename, "JSON files (*.json)")
        if not file_path:
            return
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(project_data, f, ensure_ascii=False, indent=2)
            QMessageBox.information(self, "Успех", f"Проект сохранен: {file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить проект: {e}")

    def load_project(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Загрузить проект", "", "JSON files (*.json)")
        if not file_path:
            return
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                project_data = json.load(f)
            self.load_project_data(project_data)
            QMessageBox.information(self, "Успех", "Проект загружен")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить проект: {e}")

    def load_project_data(self, project_data: Dict[str, Any]):
        self.project_name_edit.setText(project_data.get("project_name", ""))
        self.project_code_edit.setText(project_data.get("project_code", ""))
        self.client_name_edit.setText(project_data.get("client_name", ""))
        self.contact_person_edit.setText(project_data.get("contact_person", ""))
        self.email_edit.setText(project_data.get("email", ""))

        for w in self.language_pairs.values():
            w.setParent(None)
        self.language_pairs.clear()
        self.pair_headers.clear()

        for pair_data in project_data.get("language_pairs", []):
            pair_key = pair_data["pair_name"]  # в твоём текущем формате это строка, оставляем как есть
            header_title = pair_data.get("header_title", pair_key)
            widget = LanguagePairWidget(pair_key)
            self.language_pairs[pair_key] = widget
            self.pairs_layout.addWidget(widget)
            self.pair_headers[pair_key] = header_title

            services = pair_data.get("services", {})
            if "translation" in services:
                widget.translation_group.setChecked(True)
                self.load_table_data(widget.translation_group.table, services["translation"])

        self.update_pairs_list()

        additional = project_data.get("additional_services", {})
        for name, data in additional.items():
            group = self.additional_services_widget.service_groups.get(name)
            if group:
                group.setChecked(True)
                self.load_table_data(group.table, data)

    def load_table_data(self, table, data: List[Dict[str, Any]]):
        for row, row_data in enumerate(data):
            if row < table.rowCount():
                table.item(row, 1).setText(str(row_data.get("volume", 0)))
                table.item(row, 2).setText(str(row_data.get("rate", 0)))
                table.item(row, 3).setText(str(row_data.get("total", 0)))
