from typing import Dict, List, Any
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QGroupBox, QTableWidget, QTableWidgetItem, QLabel, QHeaderView
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from logic.service_config import ServiceConfig

class LanguagePairWidget(QWidget):
    """Виджет для одной языковой пары (только Перевод)"""

    def __init__(self, pair_name: str):
        super().__init__()
        self.pair_name = pair_name
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout()

        title = QLabel(f"Языковая пара: {self.pair_name}")
        title.setFont(QFont("Arial", 10, QFont.Bold))
        layout.addWidget(title)

        self.services_layout = QVBoxLayout()

        # Только Перевод
        self.translation_group = self.create_service_group("Перевод", ServiceConfig.TRANSLATION_ROWS)
        self.services_layout.addWidget(self.translation_group)

        layout.addLayout(self.services_layout)
        self.setLayout(layout)

    def create_service_group(self, service_name: str, rows: List[Dict]) -> QGroupBox:
        group = QGroupBox(service_name)
        group.setCheckable(True)
        group.setChecked(False)

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

        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)

        group.setLayout(QVBoxLayout())
        group.layout().addWidget(table)

        setattr(group, 'table', table)
        setattr(group, 'rows_config', rows)
        setattr(group, 'base_rate_row', base_rate_row)
        return group

    def update_rates_and_sums(self, table: QTableWidget, rows: List[Dict], base_rate_row: int):
        try:
            base_rate = 0.0
            if base_rate_row is not None:
                base_rate = float(table.item(base_rate_row, 2).text() or "0")

            for row in range(table.rowCount()):
                row_cfg = rows[row]
                if not row_cfg["is_base"] and base_rate_row is not None:
                    auto_rate = base_rate * row_cfg["multiplier"]
                    table.item(row, 2).setText(f"{auto_rate:.2f}")

                volume = float(table.item(row, 1).text() or "0")
                rate = float(table.item(row, 2).text() or "0")
                table.item(row, 3).setText(f"{volume * rate:.2f}")
        except (ValueError, AttributeError):
            pass

    def get_data(self) -> Dict[str, Any]:
        data = {"pair_name": self.pair_name, "services": {}}
        if self.translation_group.isChecked():
            data["services"]["translation"] = self._get_table_data(self.translation_group.table)
        return data

    def _get_table_data(self, table: QTableWidget) -> List[Dict[str, Any]]:
        out = []
        for row in range(table.rowCount()):
            out.append({
                "parameter": table.item(row, 0).text(),
                "volume": float(table.item(row, 1).text() or "0"),
                "rate":   float(table.item(row, 2).text() or "0"),
                "total":  float(table.item(row, 3).text() or "0"),
            })
        return out
