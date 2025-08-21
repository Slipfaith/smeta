from typing import List, Dict, Any

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QGroupBox, QTableWidget, QTableWidgetItem,
    QLabel, QHeaderView, QSizePolicy, QMenu
)
from PySide6.QtCore import Qt


class ProjectSetupWidget(QWidget):
    """Table for project setup and management costs."""

    def __init__(self, initial_volume: float = 0.0):
        super().__init__()
        self._setup_ui(initial_volume)

    def _setup_ui(self, initial_volume: float):
        layout = QVBoxLayout()

        group = QGroupBox("Запуск и управление проектом")
        group.setCheckable(False)
        vbox = QVBoxLayout()

        self.table = QTableWidget(1, 4)
        self.table.setHorizontalHeaderLabels([
            "Параметр", "Объем", "Ставка (руб)", "Сумма (руб)"
        ])
        self.table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.table.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self.table.setWordWrap(False)

        self.table.setItem(0, 0, QTableWidgetItem("Запуск и управление проектом"))
        self.table.setItem(0, 1, QTableWidgetItem(str(initial_volume)))
        self.table.setItem(0, 2, QTableWidgetItem("0.00"))
        total_item = QTableWidgetItem("0.00")
        total_item.setFlags(Qt.ItemIsEnabled)
        self.table.setItem(0, 3, total_item)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)

        self.table.itemChanged.connect(self.update_sums)

        vbox.addWidget(self.table)

        # context menu for adding/removing rows
        def show_menu(pos):
            row = self.table.rowAt(pos.y())
            if row < 0:
                row = self.table.rowCount() - 1
            menu = QMenu(self.table)
            add_act = menu.addAction("Добавить строку")
            del_act = menu.addAction("Удалить строку")
            action = menu.exec(self.table.mapToGlobal(pos))
            if action == add_act:
                self.add_row_after(row)
            elif action == del_act:
                self.remove_row_at(row)

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(show_menu)

        self.subtotal_label = QLabel("Промежуточная сумма: 0.00 ₽")
        self.subtotal_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        vbox.addWidget(self.subtotal_label)

        group.setLayout(vbox)
        layout.addWidget(group)
        self.setLayout(layout)

        self._fit_table_height(self.table)
        self.update_sums()

    # ---------- helpers ----------
    def _fit_table_height(self, table: QTableWidget):
        header_h = table.horizontalHeader().height()
        rows_h = sum(table.rowHeight(r) for r in range(table.rowCount()))
        frame = table.frameWidth() * 2
        total = header_h + rows_h + frame + 2
        table.setMinimumHeight(total)
        table.setMaximumHeight(total)

    # ---------- data operations ----------
    def add_row(self):
        self.add_row_after(self.table.rowCount() - 1)

    def add_row_after(self, row: int):
        insert_at = row + 1
        self.table.insertRow(insert_at)
        self.table.setItem(insert_at, 0, QTableWidgetItem("Новая строка"))
        self.table.setItem(insert_at, 1, QTableWidgetItem("0"))
        self.table.setItem(insert_at, 2, QTableWidgetItem("0.00"))
        total_item = QTableWidgetItem("0.00")
        total_item.setFlags(Qt.ItemIsEnabled)
        self.table.setItem(insert_at, 3, total_item)
        self._fit_table_height(self.table)
        self.update_sums()

    def remove_row(self):
        self.remove_row_at(self.table.currentRow())

    def remove_row_at(self, row: int):
        if row >= 0:
            self.table.removeRow(row)
            self._fit_table_height(self.table)
            self.update_sums()

    def set_volume(self, value: float):
        if self.table.rowCount() == 0:
            self.add_row()
        item = self.table.item(0, 1)
        self.table.blockSignals(True)
        item.setText(f"{value}")
        self.table.blockSignals(False)
        self.update_sums()

    # ---------- calculations ----------
    def update_sums(self):
        try:
            subtotal = 0.0
            for row in range(self.table.rowCount()):
                volume = float((self.table.item(row, 1).text() if self.table.item(row, 1) else "0") or "0")
                rate = float((self.table.item(row, 2).text() if self.table.item(row, 2) else "0") or "0")
                total = volume * rate
                total_item = self.table.item(row, 3)
                if total_item is None:
                    total_item = QTableWidgetItem("0.00")
                    total_item.setFlags(Qt.ItemIsEnabled)
                    self.table.setItem(row, 3, total_item)
                self.table.blockSignals(True)
                total_item.setText(f"{total:.2f}")
                self.table.blockSignals(False)
                subtotal += total
            self.subtotal_label.setText(f"Промежуточная сумма: {subtotal:.2f} ₽")
            self._fit_table_height(self.table)
        except Exception:
            pass

    # ---------- accessors ----------
    def get_data(self) -> List[Dict[str, Any]]:
        data: List[Dict[str, Any]] = []
        for row in range(self.table.rowCount()):
            data.append({
                "parameter": self.table.item(row, 0).text() if self.table.item(row, 0) else "",
                "volume": float((self.table.item(row, 1).text() if self.table.item(row, 1) else "0") or "0"),
                "rate": float((self.table.item(row, 2).text() if self.table.item(row, 2) else "0") or "0"),
                "total": float((self.table.item(row, 3).text() if self.table.item(row, 3) else "0") or "0"),
            })
        return data

    def load_data(self, rows: List[Dict[str, Any]]):
        self.table.blockSignals(True)
        self.table.setRowCount(len(rows))
        for i, row_data in enumerate(rows):
            self.table.setItem(i, 0, QTableWidgetItem(row_data.get("parameter", "")))
            self.table.setItem(i, 1, QTableWidgetItem(str(row_data.get("volume", 0))))
            self.table.setItem(i, 2, QTableWidgetItem(str(row_data.get("rate", 0))))
            total_item = QTableWidgetItem(str(row_data.get("total", 0)))
            total_item.setFlags(Qt.ItemIsEnabled)
            self.table.setItem(i, 3, total_item)
        self.table.blockSignals(False)
        self._fit_table_height(self.table)
        self.update_sums()
