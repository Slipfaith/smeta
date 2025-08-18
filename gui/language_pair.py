from typing import Dict, List, Any
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QGroupBox, QTableWidget, QTableWidgetItem, QLabel, QHeaderView
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from logic.service_config import ServiceConfig

class LanguagePairWidget(QWidget):
    """Виджет для одной языковой пары"""

    def __init__(self, pair_name: str):
        super().__init__()
        self.pair_name = pair_name
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout()

        # Заголовок языковой пары
        title = QLabel(f"Языковая пара: {self.pair_name}")
        title.setFont(QFont("Arial", 10, QFont.Bold))
        layout.addWidget(title)

        # Услуги для этой языковой пары
        self.services_layout = QVBoxLayout()

        # Перевод
        self.translation_group = self.create_service_group("Перевод", ServiceConfig.TRANSLATION_ROWS)
        self.services_layout.addWidget(self.translation_group)

        # Редактирование
        self.editing_group = self.create_service_group("Редактирование", ServiceConfig.EDITING_ROWS)
        self.services_layout.addWidget(self.editing_group)

        layout.addLayout(self.services_layout)
        self.setLayout(layout)

    def create_service_group(self, service_name: str, rows: List[Dict]) -> QGroupBox:
        """Создает группу для услуги с таблицей параметров"""
        group = QGroupBox(service_name)
        group.setCheckable(True)
        group.setChecked(False)

        layout = QVBoxLayout()

        # Таблица параметров
        table = QTableWidget(len(rows), 4)  # строки, колонки: Параметр, Объем, Ставка, Сумма
        table.setHorizontalHeaderLabels(["Параметр", "Объем", "Ставка (руб)", "Сумма (руб)"])

        # Сохраняем базовую ставку для автоматических расчетов
        base_rate_row = None

        for i, row_info in enumerate(rows):
            # Название параметра
            table.setItem(i, 0, QTableWidgetItem(row_info["name"]))

            # Объем
            volume_item = QTableWidgetItem("0")
            table.setItem(i, 1, volume_item)

            # Ставка
            rate_item = QTableWidgetItem("0.00")
            if not row_info["is_base"]:
                rate_item.setFlags(Qt.ItemIsEnabled)  # только чтение для автоматических ставок
            else:
                if base_rate_row is None:
                    base_rate_row = i  # запоминаем первую базовую ставку
            table.setItem(i, 2, rate_item)

            # Сумма (только чтение)
            sum_item = QTableWidgetItem("0.00")
            sum_item.setFlags(Qt.ItemIsEnabled)  # только чтение
            table.setItem(i, 3, sum_item)

        # Подключаем обновление ставок и сумм при изменении данных
        table.itemChanged.connect(lambda item: self.update_rates_and_sums(table, rows, base_rate_row))

        # Настройка ширины колонок
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)

        layout.addWidget(table)
        group.setLayout(layout)

        # Сохраняем ссылки на таблицу и конфигурацию строк
        setattr(group, 'table', table)
        setattr(group, 'rows_config', rows)
        setattr(group, 'base_rate_row', base_rate_row)

        return group

    def update_rates_and_sums(self, table: QTableWidget, rows: List[Dict], base_rate_row: int):
        """Обновляет ставки и суммы в таблице"""
        try:
            # Получаем базовую ставку
            base_rate = 0.0
            if base_rate_row is not None:
                base_rate = float(table.item(base_rate_row, 2).text() or "0")

            # Обновляем все строки
            for row in range(table.rowCount()):
                row_config = rows[row]

                # Обновляем ставку для неосновных строк
                if not row_config["is_base"] and base_rate_row is not None:
                    auto_rate = base_rate * row_config["multiplier"]
                    table.item(row, 2).setText(f"{auto_rate:.2f}")

                # Обновляем сумму
                volume = float(table.item(row, 1).text() or "0")
                rate = float(table.item(row, 2).text() or "0")
                total = volume * rate
                table.item(row, 3).setText(f"{total:.2f}")

        except (ValueError, AttributeError):
            # В случае ошибки просто пропускаем обновление
            pass

    def get_data(self) -> Dict[str, Any]:
        """Получает данные языковой пары"""
        data = {"pair_name": self.pair_name, "services": {}}

        # Перевод
        if self.translation_group.isChecked():
            data["services"]["translation"] = self.get_table_data(self.translation_group.table)

        # Редактирование
        if self.editing_group.isChecked():
            data["services"]["editing"] = self.get_table_data(self.editing_group.table)

        return data

    def get_table_data(self, table: QTableWidget) -> List[Dict[str, Any]]:
        """Получает данные из таблицы"""
        data = []
        for row in range(table.rowCount()):
            row_data = {
                "parameter": table.item(row, 0).text() if table.item(row, 0) else "",
                "volume": float(table.item(row, 1).text() or "0") if table.item(row, 1) else 0,
                "rate": float(table.item(row, 2).text() or "0") if table.item(row, 2) else 0,
                "total": float(table.item(row, 3).text() or "0") if table.item(row, 3) else 0
            }
            data.append(row_data)
        return data
