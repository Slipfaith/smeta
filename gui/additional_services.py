from typing import Dict, List, Any
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QGroupBox, QTableWidget, QTableWidgetItem, QLabel, QHeaderView
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from logic.service_config import ServiceConfig

class AdditionalServicesWidget(QWidget):
    """Виджет для дополнительных услуг"""

    def __init__(self):
        super().__init__()
        self.service_groups = {}
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout()

        title = QLabel("Дополнительные услуги")
        title.setFont(QFont("Arial", 12, QFont.Bold))
        layout.addWidget(title)

        # Создаем группы для каждой дополнительной услуги
        for service_name, rows in ServiceConfig.ADDITIONAL_SERVICES.items():
            group = self.create_service_group(service_name, rows)
            self.service_groups[service_name] = group
            layout.addWidget(group)

        self.setLayout(layout)

    def create_service_group(self, service_name: str, rows: List[Dict]) -> QGroupBox:
        """Создает группу для дополнительной услуги"""
        group = QGroupBox(service_name)
        group.setCheckable(True)
        group.setChecked(False)

        layout = QVBoxLayout()

        table = QTableWidget(len(rows), 4)
        table.setHorizontalHeaderLabels(["Параметр", "Объем", "Ставка (руб)", "Сумма (руб)"])

        base_rate_row = None

        for i, row_info in enumerate(rows):
            table.setItem(i, 0, QTableWidgetItem(row_info["name"]))
            table.setItem(i, 1, QTableWidgetItem("0"))

            rate_item = QTableWidgetItem("0.00")
            if not row_info["is_base"]:
                rate_item.setFlags(Qt.ItemIsEnabled)
            else:
                if base_rate_row is None:
                    base_rate_row = i
            table.setItem(i, 2, rate_item)

            sum_item = QTableWidgetItem("0.00")
            sum_item.setFlags(Qt.ItemIsEnabled)
            table.setItem(i, 3, sum_item)

        table.itemChanged.connect(lambda item: self.update_rates_and_sums(table, rows, base_rate_row))

        # Настройка ширины колонок
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)

        layout.addWidget(table)
        group.setLayout(layout)

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
            pass

    def get_data(self) -> Dict[str, Any]:
        """Получает данные дополнительных услуг"""
        data = {}
        for service_name, group in self.service_groups.items():
            if group.isChecked():
                data[service_name] = self.get_table_data(group.table)
        return data

    def get_table_data(self, table: QTableWidget) -> List[Dict[str, Any]]:
        """Получает данные из таблицы"""
        data = []
        for row in range(table.rowCount()):
            row_data = {
                "parameter": table.item(row, 0).text(),
                "volume": float(table.item(row, 1).text() or "0"),
                "rate": float(table.item(row, 2).text() or "0"),
                "total": float(table.item(row, 3).text() or "0")
            }
            data.append(row_data)
        return data
