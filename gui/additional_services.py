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
    QAbstractItemView,
    QDoubleSpinBox,
)
from .utils import format_rate, _to_float, format_amount
from logic.translation_config import tr


class AdditionalServiceTable(QWidget):
    """One editable table of additional services."""

    remove_requested = Signal()
    subtotal_changed = Signal(float)

    def __init__(self, title: str = "Дополнительные услуги", currency_symbol: str = "₽", currency_code: str = "RUB", lang: str = "ru") -> None:
        super().__init__()
        self.currency_symbol = currency_symbol
        self.currency_code = currency_code
        self.lang = lang
        self._subtotal = 0.0
        self._discount_percent = 0.0
        self._markup_percent = 0.0
        self._setup_ui(title)

    # ------------------------------------------------------------------ UI
    def _setup_ui(self, title: str) -> None:
        layout = QVBoxLayout()

        header = QHBoxLayout()
        self.header_edit = QLineEdit(tr(title, self.lang))
        header.addWidget(self.header_edit)
        header.addStretch()
        layout.addLayout(header)

        self.table = QTableWidget(1, 5)
        symbol_suffix = f" ({self.currency_symbol})" if self.currency_symbol else ""
        self.table.setHorizontalHeaderLabels([
            tr("Параметр", self.lang),
            tr("Ед-ца", self.lang),
            tr("Объем", self.lang),
            f"{tr('Ставка', self.lang)}{symbol_suffix}",
            f"{tr('Сумма', self.lang)}{symbol_suffix}",
        ])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)

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

        discount_layout = QHBoxLayout()
        self.discount_label = QLabel(tr("Скидка, %", self.lang))
        self.discount_spin = QDoubleSpinBox()
        self.discount_spin.setRange(0, 100)
        self.discount_spin.setDecimals(1)
        self.discount_spin.setSingleStep(1.0)
        self.discount_spin.setValue(0.0)
        self.discount_spin.valueChanged.connect(self._on_discount_changed)
        discount_layout.addWidget(self.discount_label)
        discount_layout.addWidget(self.discount_spin)
        discount_layout.addStretch()
        self.discounted_label = QLabel(
            f"{tr('Сумма скидки', self.lang)}: 0.00{f' {self.currency_symbol}' if self.currency_symbol else ''}"
        )
        self.discounted_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        discount_layout.addWidget(self.discounted_label)
        layout.addLayout(discount_layout)

        markup_layout = QHBoxLayout()
        self.markup_label = QLabel(tr("Наценка, %", self.lang))
        self.markup_spin = QDoubleSpinBox()
        self.markup_spin.setRange(0, 100)
        self.markup_spin.setDecimals(1)
        self.markup_spin.setSingleStep(1.0)
        self.markup_spin.setValue(0.0)
        self.markup_spin.valueChanged.connect(self._on_markup_changed)
        markup_layout.addWidget(self.markup_label)
        markup_layout.addWidget(self.markup_spin)
        markup_layout.addStretch()
        self.markup_amount_label = QLabel(
            f"{tr('Сумма наценки', self.lang)}: 0.00{f' {self.currency_symbol}' if self.currency_symbol else ''}"
        )
        self.markup_amount_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        markup_layout.addWidget(self.markup_amount_label)
        layout.addLayout(markup_layout)

        subtotal_suffix = f" {self.currency_symbol}" if self.currency_symbol else ""
        self.subtotal_label = QLabel(
            f"{tr('Промежуточная сумма', self.lang)}: 0.00{subtotal_suffix}"
        )
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
        del_selected_act = menu.addAction(tr("Удалить выбранные строки", self.lang))
        if self.table.rowCount() <= 1:
            del_act.setEnabled(False)
            del_selected_act.setEnabled(False)
        selected_rows = {
            index.row() for index in self.table.selectedIndexes()
        }
        if len(selected_rows) <= 1:
            del_selected_act.setEnabled(False)
        action = menu.exec(self.table.mapToGlobal(pos))
        if action == add_act:
            self.add_row_after(row)
        elif action == del_act:
            self.remove_row(row)
        elif action == del_selected_act:
            self.remove_selected_rows()

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

    def remove_selected_rows(self) -> None:
        rows = sorted({index.row() for index in self.table.selectedIndexes()})
        if not rows:
            return
        remaining = self.table.rowCount()
        removable = sorted(r for r in rows if 0 <= r < remaining)
        if not removable:
            return
        max_removable = max(0, remaining - 1)
        if max_removable == 0:
            return
        if len(removable) > max_removable:
            removable = removable[-max_removable:]
        for row in sorted(removable, reverse=True):
            self.table.removeRow(row)
        if self.table.rowCount() == 0:
            self.add_row_after(-1)
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
            item.setText(format_amount(total, self.lang))

        suffix = f" {self.currency_symbol}" if self.currency_symbol else ""
        self.subtotal_label.setText(
            f"{tr('Промежуточная сумма', self.lang)}: {format_amount(subtotal, self.lang)}{suffix}"
        )
        self._subtotal = subtotal
        self._update_discount_label()
        self.subtotal_changed.emit(self.get_subtotal())

    def _text(self, row: int, col: int) -> str:
        item = self.table.item(row, col)
        return item.text() if item else "0"

    def get_subtotal(self) -> float:
        base = self._subtotal
        discount_amount = base * (self._discount_percent / 100.0)
        markup_amount = base * (self._markup_percent / 100.0)
        return base - discount_amount + markup_amount

    def get_discount_amount(self) -> float:
        return self._subtotal * (self._discount_percent / 100.0)

    def get_markup_amount(self) -> float:
        return self._subtotal * (self._markup_percent / 100.0)

    def _update_discount_label(self) -> None:
        suffix = f" {self.currency_symbol}" if self.currency_symbol else ""
        if hasattr(self, "discount_label"):
            self.discount_label.setText(tr("Скидка, %", self.lang))
        if hasattr(self, "markup_label"):
            self.markup_label.setText(tr("Наценка, %", self.lang))
        discount_amount = self._subtotal * (self._discount_percent / 100.0)
        markup_amount = self._subtotal * (self._markup_percent / 100.0)
        effective_total = self.get_subtotal()
        if hasattr(self, "discounted_label"):
            self.discounted_label.setText(
                f"{tr('Сумма скидки', self.lang)}: {format_amount(discount_amount, self.lang)}{suffix}"
            )
        if hasattr(self, "markup_amount_label"):
            self.markup_amount_label.setText(
                f"{tr('Сумма наценки', self.lang)}: {format_amount(markup_amount, self.lang)}{suffix}"
            )
        if hasattr(self, "subtotal_label"):
            self.subtotal_label.setText(
                f"{tr('Промежуточная сумма', self.lang)}: {format_amount(effective_total, self.lang)}{suffix}"
            )

    def _on_discount_changed(self, value: float) -> None:
        self._discount_percent = max(0.0, min(100.0, float(value)))
        self._update_discount_label()
        self.subtotal_changed.emit(self.get_subtotal())

    def _on_markup_changed(self, value: float) -> None:
        self._markup_percent = max(0.0, min(100.0, float(value)))
        self._update_discount_label()
        self.subtotal_changed.emit(self.get_subtotal())

    def get_discount_percent(self) -> float:
        return self._discount_percent

    def set_discount_percent(self, value: float) -> None:
        self._discount_percent = max(0.0, min(100.0, float(value)))
        if hasattr(self, "discount_spin"):
            self.discount_spin.blockSignals(True)
            self.discount_spin.setValue(self._discount_percent)
            self.discount_spin.blockSignals(False)
        self._update_discount_label()
        self.subtotal_changed.emit(self.get_subtotal())

    def get_markup_percent(self) -> float:
        return self._markup_percent

    def set_markup_percent(self, value: float) -> None:
        self._markup_percent = max(0.0, min(100.0, float(value)))
        if hasattr(self, "markup_spin"):
            self.markup_spin.blockSignals(True)
            self.markup_spin.setValue(self._markup_percent)
            self.markup_spin.blockSignals(False)
        self._update_discount_label()
        self.subtotal_changed.emit(self.get_subtotal())

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
            "discount_percent": self.get_discount_percent(),
            "discount_amount": self.get_discount_amount(),
            "markup_percent": self.get_markup_percent(),
            "markup_amount": self.get_markup_amount(),
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
        self.set_discount_percent(data.get("discount_percent", 0.0))
        self.set_markup_percent(data.get("markup_percent", 0.0))

    def set_currency(self, symbol: str, code: str) -> None:
        self.currency_symbol = symbol
        self.currency_code = code
        symbol_suffix = f" ({symbol})" if symbol else ""
        self.table.setHorizontalHeaderLabels([
            tr("Параметр", self.lang),
            tr("Ед-ца", self.lang),
            tr("Объем", self.lang),
            f"{tr('Ставка', self.lang)}{symbol_suffix}",
            f"{tr('Сумма', self.lang)}{symbol_suffix}",
        ])
        self.update_sums()

    def convert_rates(self, multiplier: float) -> None:
        """Multiply all rate values by *multiplier* and update totals."""
        for r in range(self.table.rowCount()):
            item = self.table.item(r, 3)
            if item is None:
                continue
            rate = _to_float(item.text())
            sep = '.' if self.lang == 'en' else ','
            item.setText(format_rate(rate * multiplier, sep))
        self.update_sums()

    def set_language(self, lang: str) -> None:
        self.lang = lang
        self.header_edit.setText(tr("Дополнительные услуги", lang))
        if hasattr(self, "discount_label"):
            self.discount_label.setText(tr("Скидка, %", lang))
        if hasattr(self, "discounted_label"):
            suffix = f" {self.currency_symbol}" if self.currency_symbol else ""
            self.discounted_label.setText(
                f"{tr('Сумма скидки', lang)}: 0.00{suffix}"
            )
        if hasattr(self, "markup_label"):
            self.markup_label.setText(tr("Наценка, %", lang))
        if hasattr(self, "markup_amount_label"):
            suffix = f" {self.currency_symbol}" if self.currency_symbol else ""
            self.markup_amount_label.setText(
                f"{tr('Сумма наценки', lang)}: 0.00{suffix}"
            )
        self.set_currency(self.currency_symbol, self.currency_code)
        self.update_sums()


class AdditionalServicesWidget(QWidget):
    """Container managing multiple additional service tables."""

    subtotal_changed = Signal(float)

    def __init__(self, currency_symbol: str = "₽", currency_code: str = "RUB", lang: str = "ru") -> None:
        super().__init__()
        self.currency_symbol = currency_symbol
        self.currency_code = currency_code
        self.lang = lang
        self.tables: List[AdditionalServiceTable] = []
        self._subtotal = 0.0
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
        table.subtotal_changed.connect(self._emit_subtotal)
        self.tables.append(table)
        self.tables_layout.addWidget(table)
        if data:
            table.load_data(data)
        self._emit_subtotal()

    def remove_table(self, table: AdditionalServiceTable) -> None:
        if table in self.tables and len(self.tables) > 1:
            self.tables.remove(table)
            table.setParent(None)
            self._emit_subtotal()

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
        self._emit_subtotal()

    def convert_rates(self, multiplier: float) -> None:
        """Multiply all rate values by *multiplier* across all tables."""
        for tbl in self.tables:
            tbl.convert_rates(multiplier)
        self._emit_subtotal()

    def set_language(self, lang: str) -> None:
        self.lang = lang
        self.add_btn.setText(tr("Добавить таблицу", lang))
        for tbl in self.tables:
            tbl.set_language(lang)
        self._emit_subtotal()

    def _emit_subtotal(self) -> None:
        total = sum(tbl.get_subtotal() for tbl in self.tables)
        self._subtotal = total
        self.subtotal_changed.emit(total)

    def get_subtotal(self) -> float:
        return self._subtotal

    def get_discount_amount(self) -> float:
        return sum(tbl.get_discount_amount() for tbl in self.tables)

    def get_markup_amount(self) -> float:
        return sum(tbl.get_markup_amount() for tbl in self.tables)

