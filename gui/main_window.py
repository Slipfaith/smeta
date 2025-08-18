import json
from datetime import datetime
from typing import Dict, List, Any

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QGroupBox, QTextEdit, QFileDialog, QMessageBox, QScrollArea, QTabWidget, QSplitter, QFrame
)
from PySide6.QtCore import Qt

from gui.language_pair import LanguagePairWidget
from gui.additional_services import AdditionalServicesWidget
from gui.styles import APP_STYLE
from logic.excel_exporter import ExcelExporter

class TranslationCostCalculator(QMainWindow):
    """Главное окно приложения"""

    def __init__(self):
        super().__init__()
        self.language_pairs = {}  # словарь виджетов языковых пар
        self.setup_ui()
        self.setup_style()

    def setup_ui(self):
        self.setWindowTitle("Калькулятор стоимости переводческих проектов")
        self.setGeometry(100, 100, 1200, 800)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QHBoxLayout()

        # Создаем Splitter для изменяемого размера панелей
        splitter = QSplitter(Qt.Horizontal)

        # Левая панель - основные данные и управление
        left_panel = self.create_left_panel()
        splitter.addWidget(left_panel)

        # Правая панель - языковые пары и услуги
        right_panel = self.create_right_panel()
        splitter.addWidget(right_panel)

        # Устанавливаем пропорции
        splitter.setSizes([300, 900])

        main_layout.addWidget(splitter)
        central_widget.setLayout(main_layout)

    def create_left_panel(self) -> QWidget:
        """Создает левую панель с основными данными"""
        widget = QWidget()
        layout = QVBoxLayout()

        # Основная информация о проекте
        project_group = QGroupBox("Информация о проекте")
        project_layout = QVBoxLayout()

        # Название проекта
        project_layout.addWidget(QLabel("Название проекта:"))
        self.project_name_edit = QLineEdit()
        project_layout.addWidget(self.project_name_edit)

        # Номер проекта
        project_layout.addWidget(QLabel("Номер/код проекта:"))
        self.project_code_edit = QLineEdit()
        project_layout.addWidget(self.project_code_edit)

        # Клиент
        project_layout.addWidget(QLabel("Название клиента:"))
        self.client_name_edit = QLineEdit()
        project_layout.addWidget(self.client_name_edit)

        # Контактное лицо
        project_layout.addWidget(QLabel("Контактное лицо:"))
        self.contact_person_edit = QLineEdit()
        project_layout.addWidget(self.contact_person_edit)

        # Email
        project_layout.addWidget(QLabel("E-mail:"))
        self.email_edit = QLineEdit()
        project_layout.addWidget(self.email_edit)

        project_group.setLayout(project_layout)
        layout.addWidget(project_group)

        # Управление языковыми парами
        pairs_group = QGroupBox("Языковые пары")
        pairs_layout = QVBoxLayout()

        # Поля для добавления новой пары
        add_pair_layout = QHBoxLayout()

        self.source_lang_edit = QLineEdit()
        self.source_lang_edit.setPlaceholderText("Исходный язык")
        add_pair_layout.addWidget(self.source_lang_edit)

        add_pair_layout.addWidget(QLabel("→"))

        self.target_lang_edit = QLineEdit()
        self.target_lang_edit.setPlaceholderText("Целевой язык")
        add_pair_layout.addWidget(self.target_lang_edit)

        pairs_layout.addLayout(add_pair_layout)

        # Кнопка добавления пары
        self.add_pair_btn = QPushButton("Добавить языковую пару")
        self.add_pair_btn.clicked.connect(self.add_language_pair)
        pairs_layout.addWidget(self.add_pair_btn)

        # Список текущих пар
        pairs_layout.addWidget(QLabel("Текущие пары:"))
        self.pairs_list = QTextEdit()
        self.pairs_list.setMaximumHeight(100)
        self.pairs_list.setReadOnly(True)
        pairs_layout.addWidget(self.pairs_list)

        pairs_group.setLayout(pairs_layout)
        layout.addWidget(pairs_group)

        # Кнопки действий
        actions_group = QGroupBox("Действия")
        actions_layout = QVBoxLayout()

        # Выбор шаблона
        template_layout = QHBoxLayout()
        self.template_path_edit = QLineEdit()
        self.template_path_edit.setPlaceholderText("Путь к шаблону Excel")
        template_layout.addWidget(self.template_path_edit)

        select_template_btn = QPushButton("Выбрать")
        select_template_btn.clicked.connect(self.select_template)
        template_layout.addWidget(select_template_btn)

        actions_layout.addLayout(template_layout)

        # Сохранение
        self.save_excel_btn = QPushButton("Сохранить Excel")
        self.save_excel_btn.clicked.connect(self.save_excel)
        actions_layout.addWidget(self.save_excel_btn)

        self.save_pdf_btn = QPushButton("Сохранить PDF")
        self.save_pdf_btn.clicked.connect(self.save_pdf)
        actions_layout.addWidget(self.save_pdf_btn)

        # Загрузка/сохранение проекта
        actions_layout.addWidget(QFrame())  # разделитель

        self.save_project_btn = QPushButton("Сохранить проект")
        self.save_project_btn.clicked.connect(self.save_project)
        actions_layout.addWidget(self.save_project_btn)

        self.load_project_btn = QPushButton("Загрузить проект")
        self.load_project_btn.clicked.connect(self.load_project)
        actions_layout.addWidget(self.load_project_btn)

        actions_group.setLayout(actions_layout)
        layout.addWidget(actions_group)

        layout.addStretch()  # Добавляем растяжение внизу
        widget.setLayout(layout)
        return widget

    def create_right_panel(self) -> QWidget:
        """Создает правую панель с услугами"""
        widget = QWidget()
        layout = QVBoxLayout()

        # Создаем вкладки
        self.tabs = QTabWidget()

        # Вкладка для языковых пар
        self.pairs_scroll = QScrollArea()
        self.pairs_widget = QWidget()
        self.pairs_layout = QVBoxLayout()
        self.pairs_widget.setLayout(self.pairs_layout)
        self.pairs_scroll.setWidget(self.pairs_widget)
        self.pairs_scroll.setWidgetResizable(True)

        self.tabs.addTab(self.pairs_scroll, "Языковые пары")

        # Вкладка для дополнительных услуг
        self.additional_services_widget = AdditionalServicesWidget()
        additional_scroll = QScrollArea()
        additional_scroll.setWidget(self.additional_services_widget)
        additional_scroll.setWidgetResizable(True)

        self.tabs.addTab(additional_scroll, "Дополнительные услуги")

        layout.addWidget(self.tabs)
        widget.setLayout(layout)
        return widget

    def setup_style(self):
        """Настраивает стили приложения"""
        self.setStyleSheet(APP_STYLE)

    def add_language_pair(self):
        """Добавляет новую языковую пару"""
        source_lang = self.source_lang_edit.text().strip()
        target_lang = self.target_lang_edit.text().strip()

        if not source_lang or not target_lang:
            QMessageBox.warning(self, "Ошибка", "Введите оба языка")
            return

        pair_name = f"{source_lang} → {target_lang}"

        if pair_name in self.language_pairs:
            QMessageBox.warning(self, "Ошибка", "Такая языковая пара уже существует")
            return

        # Создаем виджет для языковой пары
        pair_widget = LanguagePairWidget(pair_name)
        self.language_pairs[pair_name] = pair_widget

        # Добавляем в интерфейс
        self.pairs_layout.addWidget(pair_widget)

        # Обновляем список пар
        self.update_pairs_list()

        # Очищаем поля ввода
        self.source_lang_edit.clear()
        self.target_lang_edit.clear()

    def update_pairs_list(self):
        """Обновляет список языковых пар"""
        pairs_text = "\n".join(self.language_pairs.keys())
        self.pairs_list.setText(pairs_text)

    def select_template(self):
        """Выбирает файл шаблона Excel"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выбрать шаблон Excel", "",
            "Excel files (*.xlsx *.xls)"
        )
        if file_path:
            self.template_path_edit.setText(file_path)

    def collect_project_data(self) -> Dict[str, Any]:
        """Собирает все данные проекта"""
        data = {
            "project_name": self.project_name_edit.text(),
            "project_code": self.project_code_edit.text(),
            "client_name": self.client_name_edit.text(),
            "contact_person": self.contact_person_edit.text(),
            "email": self.email_edit.text(),
            "language_pairs": [],
            "additional_services": {}
        }

        # Собираем данные языковых пар
        for pair_widget in self.language_pairs.values():
            pair_data = pair_widget.get_data()
            if pair_data["services"]:  # Только если есть выбранные услуги
                data["language_pairs"].append(pair_data)

        # Собираем данные дополнительных услуг
        additional_data = self.additional_services_widget.get_data()
        if additional_data:
            data["additional_services"] = additional_data

        return data

    def save_excel(self):
        """Сохраняет данные в Excel файл"""
        if not self.client_name_edit.text().strip():
            QMessageBox.warning(self, "Ошибка", "Введите название клиента")
            return

        project_data = self.collect_project_data()

        if not project_data["language_pairs"] and not project_data["additional_services"]:
            QMessageBox.warning(self, "Ошибка", "Добавьте хотя бы одну услугу")
            return

        # Генерируем имя файла
        client_name = project_data["client_name"].replace(" ", "_")
        date_str = datetime.now().strftime("%Y%m%d")
        filename = f"КП_{client_name}_{date_str}.xlsx"

        # Выбираем папку для сохранения
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить Excel файл", filename,
            "Excel files (*.xlsx)"
        )

        if not file_path:
            return

        # Экспортируем данные
        template_path = self.template_path_edit.text().strip()
        exporter = ExcelExporter(template_path if template_path else None)

        if exporter.export_to_excel(project_data, file_path):
            QMessageBox.information(self, "Успех", f"Файл сохранен: {file_path}")
        else:
            QMessageBox.critical(self, "Ошибка", "Не удалось сохранить файл")

    def save_pdf(self):
        """Сохраняет данные в PDF файл"""
        QMessageBox.information(
            self,
            "Функция PDF",
            "Функция сохранения в PDF будет реализована в следующей версии.\n"
            "Пока вы можете сохранить Excel файл и конвертировать его в PDF вручную."
        )

    def save_project(self):
        """Сохраняет проект в JSON файл"""
        if not self.project_name_edit.text().strip():
            QMessageBox.warning(self, "Ошибка", "Введите название проекта")
            return

        project_data = self.collect_project_data()

        # Генерируем имя файла
        project_name = project_data["project_name"].replace(" ", "_")
        filename = f"Проект_{project_name}.json"

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить проект", filename,
            "JSON files (*.json)"
        )

        if not file_path:
            return

        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(project_data, f, ensure_ascii=False, indent=2)
            QMessageBox.information(self, "Успех", f"Проект сохранен: {file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить проект: {e}")

    def load_project(self):
        """Загружает проект из JSON файла"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Загрузить проект", "",
            "JSON files (*.json)"
        )

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
        """Загружает данные проекта в интерфейс"""
        # Основная информация
        self.project_name_edit.setText(project_data.get("project_name", ""))
        self.project_code_edit.setText(project_data.get("project_code", ""))
        self.client_name_edit.setText(project_data.get("client_name", ""))
        self.contact_person_edit.setText(project_data.get("contact_person", ""))
        self.email_edit.setText(project_data.get("email", ""))

        # Очищаем текущие языковые пары
        for pair_widget in self.language_pairs.values():
            pair_widget.setParent(None)
        self.language_pairs.clear()

        # Загружаем языковые пары
        for pair_data in project_data.get("language_pairs", []):
            pair_name = pair_data["pair_name"]

            # Создаем виджет пары
            pair_widget = LanguagePairWidget(pair_name)
            self.language_pairs[pair_name] = pair_widget
            self.pairs_layout.addWidget(pair_widget)

            # Загружаем данные услуг
            services = pair_data.get("services", {})

            if "translation" in services:
                pair_widget.translation_group.setChecked(True)
                self.load_table_data(pair_widget.translation_group.table, services["translation"])

            if "editing" in services:
                pair_widget.editing_group.setChecked(True)
                self.load_table_data(pair_widget.editing_group.table, services["editing"])

        # Обновляем список пар
        self.update_pairs_list()

        # Загружаем дополнительные услуги
        additional_services = project_data.get("additional_services", {})
        for service_name, service_data in additional_services.items():
            group = self.additional_services_widget.service_groups.get(service_name)
            if group:
                group.setChecked(True)
                self.load_table_data(group.table, service_data)

    def load_table_data(self, table, data: List[Dict[str, Any]]):
        """Загружает данные в таблицу"""
        for row, row_data in enumerate(data):
            if row < table.rowCount():
                table.item(row, 1).setText(str(row_data.get("volume", 0)))
                table.item(row, 2).setText(str(row_data.get("rate", 0)))
                table.item(row, 3).setText(str(row_data.get("total", 0)))

    def calculate_total_cost(self) -> float:
        """Вычисляет общую стоимость проекта"""
        total = 0.0

        # Суммируем языковые пары
        for pair_widget in self.language_pairs.values():
            pair_data = pair_widget.get_data()
            for service_data in pair_data.get("services", {}).values():
                for row_data in service_data:
                    total += row_data["total"]

        # Суммируем дополнительные услуги
        additional_data = self.additional_services_widget.get_data()
        for service_data in additional_data.values():
            for row_data in service_data:
                total += row_data["total"]

        return total
