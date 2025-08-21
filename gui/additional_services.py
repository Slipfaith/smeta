from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QLineEdit,
    QTableWidget,
    QTableWidgetItem,
    QLabel,
    QHeaderView,
    QMenu,
)


def _to_float(value: str) -> float:
    """Safely convert string to float."""
    try:
        return float((value or "0").replace(",", "."))
    except ValueError:
        return 0.0


class AdditionalServicesWidget(QWidget):
    """Single table for user defined additional services."""

    def __init__(self) -> None:
        super().__init__()
        self._setup_ui()

    # ------------------------------------------------------------------ UI
    def _setup_ui(self) -> None:
        layout = QVBoxLayout()

        self.header_edit = QLineEdit("Дополнительные услуги")
        layout.addWidget(self.header_edit)

        self.table = QTableWidget(1, 5)
        self.table.setHorizontalHeaderLabels([
            "Параметр",
            "Ед-ца",
            "Объем",
            "Ставка (руб)",
            "Сумма (руб)",
        ])

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)

        # initial row
        for col, text in enumerate(["", "", "0", "0.000", "0.00"]):
            item = QTableWidgetItem(text)
            if col == 4:
                item.setFlags(Qt.ItemIsEnabled)
            self.table.setItem(0, col, item)

        self.table.itemChanged.connect(self.update_sums)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_menu)

        layout.addWidget(self.table)

        self.subtotal_label = QLabel("Промежуточная сумма: 0.00 ₽")
        self.subtotal_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        layout.addWidget(self.subtotal_label)

        self.setLayout(layout)
        self.update_sums()

    # ----------------------------------------------------------------- menu
    def _show_menu(self, pos) -> None:
        row = self.table.rowAt(pos.y())
        if row < 0:
            row = self.table.rowCount() - 1
        menu = QMenu(self.table)
        add_act = menu.addAction("Добавить строку")
        del_act = menu.addAction("Удалить строку")
        if self.table.rowCount() <= 1:
            del_act.setEnabled(False)
        action = menu.exec(self.table.mapToGlobal(pos))
        if action == add_act:
            self.add_row_after(row)
        elif action == del_act:
            self.remove_row(row)

    def add_row_after(self, row: int) -> None:
        insert_at = row + 1
        self.table.insertRow(insert_at)
        for col, text in enumerate(["", "", "0", "0.000", "0.00"]):
            item = QTableWidgetItem(text)
            if col == 4:
                item.setFlags(Qt.ItemIsEnabled)
            self.table.setItem(insert_at, col, item)
        self.update_sums()

    def remove_row(self, row: int) -> None:
        if self.table.rowCount() > 1 and row >= 0:
            self.table.removeRow(row)
            self.update_sums()

    # ------------------------------------------------------------ calculations
    def update_sums(self) -> None:
        subtotal = 0.0
        for r in range(self.table.rowCount()):
            volume = _to_float(self._text(r, 2))
            rate = _to_float(self._text(r, 3))
            total = volume * rate
            subtotal += total
            item = self.table.item(r, 4)
            if item is None:
                item = QTableWidgetItem()
                item.setFlags(Qt.ItemIsEnabled)
                self.table.setItem(r, 4, item)
            item.setText(f"{total:.2f}")

        self.subtotal_label.setText(f"Промежуточная сумма: {subtotal:.2f} ₽")

    def _text(self, row: int, col: int) -> str:
        item = self.table.item(row, col)
        return item.text() if item else "0"

    # --------------------------------------------------------------- data i/o
    def get_data(self) -> dict:
        """Return data for export."""
        rows = []
        for r in range(self.table.rowCount()):
            rows.append({
                "parameter": self._text(r, 0),
                "unit": self._text(r, 1),
                "volume": _to_float(self._text(r, 2)),
                "rate": _to_float(self._text(r, 3)),
            })

        if not any(row["parameter"] or row["volume"] or row["rate"] for row in rows):
            return {}

        return {
            "header_title": self.header_edit.text(),
            "rows": rows,
        }

    def load_data(self, data: dict) -> None:
        """Load previously saved data."""
        self.header_edit.setText(data.get("header_title", ""))
        rows = data.get("rows", [])
        if not rows:
            return
        self.table.setRowCount(len(rows))
        for r, row in enumerate(rows):
            for col, key in enumerate(["parameter", "unit", "volume", "rate"]):
                val = row.get(key, "0" if col >= 2 else "")
                item = QTableWidgetItem(str(val))
                if col == 3:
                    item.setText(f"{_to_float(val):.3f}")
                self.table.setItem(r, col, item)
            total_item = QTableWidgetItem("0.00")
            total_item.setFlags(Qt.ItemIsEnabled)
            self.table.setItem(r, 4, total_item)
        self.update_sums()

