from typing import List, Dict

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QLineEdit,
    QTableWidget,
    QTableWidgetItem,
    QLabel,
    QHeaderView,
    QMenu,
    QHBoxLayout,
    QPushButton,
    QStyle,
)
from .utils import format_rate, _to_float
from logic.translation_config import tr


class AdditionalServiceTable(QWidget):
    """One editable table of additional services."""

    remove_requested = Signal()

    def __init__(self, title: str = "Дополнительные услуги", currency_symbol: str = "₽", currency_code: str = "RUB", lang: str = "ru") -> None:
        super().__init__()
        self.currency_symbol = currency_symbol
        self.currency_code = currency_code
        self.lang = lang
        self._setup_ui(title)

    # ------------------------------------------------------------------ UI
    def _setup_ui(self, title: str) -> None:
        layout = QVBoxLayout()

        header = QHBoxLayout()
        self.header_edit = QLineEdit(tr(title, self.lang))
        header.addWidget(self.header_edit)
        remove_btn = QPushButton()
        remove_btn.setIcon(self.style().standardIcon(QStyle.SP_TrashIcon))
        remove_btn.setFlat(True)
        remove_btn.setMaximumWidth(24)
        remove_btn.setToolTip("Удалить таблицу")
        remove_btn.setStyleSheet("background-color: transparent; border: none;")
        remove_btn.setContextMenuPolicy(Qt.NoContextMenu)
        remove_btn.clicked.connect(self.remove_requested.emit)
        header.addWidget(remove_btn)
        layout.addLayout(header)

        self.table = QTableWidget(1, 5)
        self.table.setHorizontalHeaderLabels([
            tr("Параметр", self.lang),
            tr("Ед-ца", self.lang),
            tr("Объем", self.lang),
            f"{tr('Ставка', self.lang)} ({self.currency_symbol})",
            f"{tr('Сумма', self.lang)} ({self.currency_symbol})",
        ])

        header_view = self.table.horizontalHeader()
        header_view.setSectionResizeMode(0, QHeaderView.Stretch)
        header_view.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header_view.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header_view.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header_view.setSectionResizeMode(4, QHeaderView.ResizeToContents)

        for col, text in enumerate(["", "", "0", "0.00", "0.00"]):
            item = QTableWidgetItem(text)
            if col == 4:
                item.setFlags(Qt.ItemIsEnabled)
            self.table.setItem(0, col, item)

        self.table.itemChanged.connect(self.update_sums)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_menu)

        layout.addWidget(self.table)

        self.subtotal_label = QLabel(f"{tr('Промежуточная сумма', self.lang)}: 0.00 {self.currency_symbol}")
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
        add_act = menu.addAction(tr("Добавить строку", self.lang))
        del_act = menu.addAction(tr("Удалить строку", self.lang))
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
        for col, text in enumerate(["", "", "0", "0.00", "0.00"]):
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
            rate_item = self.table.item(r, 3)
            rate_text = rate_item.text() if rate_item else "0"
            if self.lang == "en":
                sep = "."
            else:
                sep = "," if "," in rate_text else "."
            rate = _to_float(rate_text)
            self.table.blockSignals(True)
            if rate_item:
                rate_item.setText(format_rate(rate_text, sep))
            self.table.blockSignals(False)
            total = volume * rate
            subtotal += total
            item = self.table.item(r, 4)
            if item is None:
                item = QTableWidgetItem()
                item.setFlags(Qt.ItemIsEnabled)
                self.table.setItem(r, 4, item)
            item.setText(f"{total:.2f}")

        self.subtotal_label.setText(f"{tr('Промежуточная сумма', self.lang)}: {subtotal:.2f} {self.currency_symbol}")

    def _text(self, row: int, col: int) -> str:
        item = self.table.item(row, col)
        return item.text() if item else "0"

    # --------------------------------------------------------------- data i/o
    def get_data(self) -> Dict:
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

    def load_data(self, data: Dict) -> None:
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
                    sep = "." if self.lang == "en" else None
                    item.setText(format_rate(val, sep))
                self.table.setItem(r, col, item)
            total_item = QTableWidgetItem("0.00")
            total_item.setFlags(Qt.ItemIsEnabled)
            self.table.setItem(r, 4, total_item)
        self.update_sums()

    def set_currency(self, symbol: str, code: str) -> None:
        self.currency_symbol = symbol
        self.currency_code = code
        self.table.setHorizontalHeaderLabels([
            tr("Параметр", self.lang),
            tr("Ед-ца", self.lang),
            tr("Объем", self.lang),
            f"{tr('Ставка', self.lang)} ({symbol})",
            f"{tr('Сумма', self.lang)} ({symbol})",
        ])
        self.update_sums()

    def set_language(self, lang: str) -> None:
        self.lang = lang
        self.header_edit.setText(tr("Дополнительные услуги", lang))
        self.set_currency(self.currency_symbol, self.currency_code)
        self.update_sums()


class AdditionalServicesWidget(QWidget):
    """Container managing multiple additional service tables."""

    def __init__(self, currency_symbol: str = "₽", currency_code: str = "RUB", lang: str = "ru") -> None:
        super().__init__()
        self.currency_symbol = currency_symbol
        self.currency_code = currency_code
        self.lang = lang
        self.tables: List[AdditionalServiceTable] = []
        self._setup_ui()

    def _setup_ui(self) -> None:
        layout = QVBoxLayout()
        self.tables_layout = QVBoxLayout()
        layout.addLayout(self.tables_layout)

        self.add_btn = QPushButton(tr("Добавить таблицу", self.lang))
        self.add_btn.clicked.connect(self.add_table)
        layout.addWidget(self.add_btn)

        self.setLayout(layout)
        self.add_table()

    def add_table(self, data: Dict = None) -> None:
        table = AdditionalServiceTable(currency_symbol=self.currency_symbol, currency_code=self.currency_code, lang=self.lang)
        table.remove_requested.connect(lambda t=table: self.remove_table(t))
        self.tables.append(table)
        self.tables_layout.addWidget(table)
        if data:
            table.load_data(data)

    def remove_table(self, table: AdditionalServiceTable) -> None:
        if table in self.tables and len(self.tables) > 1:
            self.tables.remove(table)
            table.setParent(None)

    # --------------------------------------------------------------- data i/o
    def get_data(self) -> List[Dict]:
        data = []
        for tbl in self.tables:
            block = tbl.get_data()
            if block:
                data.append(block)
        return data

    def load_data(self, blocks: List[Dict]) -> None:
        for tbl in self.tables:
            tbl.setParent(None)
        self.tables.clear()
        if not blocks:
            self.add_table()
            return
        for block in blocks:
            self.add_table(block)

    def set_currency(self, symbol: str, code: str) -> None:
        self.currency_symbol = symbol
        self.currency_code = code
        for tbl in self.tables:
            tbl.set_currency(symbol, code)

    def set_language(self, lang: str) -> None:
        self.lang = lang
        self.add_btn.setText(tr("Добавить таблицу", lang))
        for tbl in self.tables:
            tbl.set_language(lang)

