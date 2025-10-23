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
    QSpinBox,
    QRadioButton,
    QWidgetAction,
    QButtonGroup,
    QCheckBox,
)
from gui.styles import GROUP_SECTION_MARGINS, GROUP_SECTION_SPACING
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
        self.enabled_checkbox = QCheckBox()
        self.enabled_checkbox.setChecked(True)
        self.enabled_checkbox.toggled.connect(self._on_enabled_toggled)
        header.addWidget(self.enabled_checkbox)

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

        self.rows_deleted: list[bool] = [False]
        self._set_row_deleted(0, False)

        layout.addWidget(self.table)

        subtotal_layout = QHBoxLayout()
        subtotal_layout.setContentsMargins(0, 0, 0, 0)
        self.subtotal_label = QLabel()
        self.subtotal_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        subtotal_layout.addWidget(self.subtotal_label, 1)

        self.modifiers_button = QPushButton("⚙️")
        self.modifiers_button.setFlat(True)
        self.modifiers_button.setCursor(Qt.PointingHandCursor)
        self.modifiers_button.clicked.connect(self._show_modifiers_menu)
        subtotal_layout.addWidget(self.modifiers_button)

        layout.addLayout(subtotal_layout)

        self.setLayout(layout)
        self.update_sums()
        self._on_enabled_toggled(True)

    # ----------------------------------------------------------------- menu
    def _show_menu(self, pos) -> None:
        row = self.table.rowAt(pos.y())
        if row < 0:
            row = self.table.rowCount() - 1
        menu = QMenu(self.table)
        add_act = menu.addAction(tr("Добавить строку", self.lang))
        del_act = menu.addAction(tr("Удалить", self.lang))
        restore_act = menu.addAction(tr("Восстановить", self.lang))

        selected_rows = sorted(
            {
                index.row()
                for index in self.table.selectedIndexes()
                if 0 <= index.row() < self.table.rowCount()
                and not self.rows_deleted[index.row()]
            }
        )
        selected_deleted_rows = sorted(
            {
                index.row()
                for index in self.table.selectedIndexes()
                if 0 <= index.row() < self.table.rowCount()
                and self.rows_deleted[index.row()]
            }
        )

        rows_to_delete = selected_rows if selected_rows else [row]
        if not self._can_delete_rows(rows_to_delete):
            del_act.setEnabled(False)

        can_restore = False
        if selected_deleted_rows:
            can_restore = True
        elif 0 <= row < len(self.rows_deleted) and self.rows_deleted[row]:
            can_restore = True
        if not can_restore:
            restore_act.setEnabled(False)

        action = menu.exec(self.table.mapToGlobal(pos))
        if action == add_act:
            self.add_row_after(row)
        elif action == del_act:
            self.remove_rows(rows_to_delete)
        elif action == restore_act:
            targets = selected_deleted_rows if selected_deleted_rows else [row]
            self.restore_rows(targets)

    def add_row_after(self, row: int) -> None:
        insert_at = row + 1
        self.table.insertRow(insert_at)
        for col, text in enumerate(["", "", "0", "0.00", "0.00"]):
            item = QTableWidgetItem(text)
            if col == 4:
                item.setFlags(Qt.ItemIsEnabled)
            self.table.setItem(insert_at, col, item)
        self.rows_deleted.insert(insert_at, False)
        self._set_row_deleted(insert_at, False)
        self.update_sums()

    def remove_row(self, row: int) -> None:
        self.remove_rows([row])

    def remove_rows(self, rows: list[int]) -> None:
        active_indices = [i for i, deleted in enumerate(self.rows_deleted) if not deleted]
        if len(active_indices) <= 1:
            return
        removable = sorted(
            {
                row
                for row in rows
                if 0 <= row < len(self.rows_deleted) and not self.rows_deleted[row]
            }
        )
        if not removable:
            return
        max_remove = len(active_indices) - 1
        if len(removable) > max_remove:
            removable = removable[-max_remove:]
        for row in removable:
            self._set_row_deleted(row, True)
        self.update_sums()

    def restore_rows(self, rows: list[int]) -> None:
        restored = False
        for row in sorted({r for r in rows if 0 <= r < len(self.rows_deleted) and self.rows_deleted[r]}):
            self._set_row_deleted(row, False)
            restored = True
        if restored:
            self.update_sums()

    def _can_delete_rows(self, rows: list[int]) -> bool:
        active = [i for i, deleted in enumerate(self.rows_deleted) if not deleted]
        if len(active) <= 1:
            return False
        valid = {
            r
            for r in rows
            if 0 <= r < len(self.rows_deleted) and not self.rows_deleted[r]
        }
        if not valid:
            return False
        remaining = len(active) - len(valid)
        return remaining >= 1

    def _set_row_deleted(self, row: int, deleted: bool) -> None:
        if not (0 <= row < len(self.rows_deleted)):
            return
        self.rows_deleted[row] = deleted
        self.table.blockSignals(True)
        try:
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item is None:
                    continue
                font = item.font()
                font.setStrikeOut(deleted)
                item.setFont(font)
                if deleted:
                    item.setForeground(Qt.gray)
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                    if col == 4:
                        item.setText(format_amount(0.0, self.lang))
                else:
                    item.setForeground(Qt.black)
                    flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable
                    if col != 4:
                        flags |= Qt.ItemIsEditable
                    item.setFlags(flags)
        finally:
            self.table.blockSignals(False)

    # ------------------------------------------------------------ calculations
    def update_sums(self) -> None:
        subtotal = 0.0
        for r in range(self.table.rowCount()):
            if r >= len(self.rows_deleted) or self.rows_deleted[r]:
                item = self.table.item(r, 4)
                if item is None:
                    item = QTableWidgetItem("0.00")
                    item.setFlags(Qt.ItemIsEnabled)
                    self.table.setItem(r, 4, item)
                item.setText(format_amount(0.0, self.lang))
                continue
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
        if not self.is_enabled():
            return 0.0
        base = self._subtotal
        discount_amount = base * (self._discount_percent / 100.0)
        markup_amount = base * (self._markup_percent / 100.0)
        return base - discount_amount + markup_amount

    def get_discount_amount(self) -> float:
        if not self.is_enabled():
            return 0.0
        return self._subtotal * (self._discount_percent / 100.0)

    def get_markup_amount(self) -> float:
        if not self.is_enabled():
            return 0.0
        return self._subtotal * (self._markup_percent / 100.0)

    def _update_discount_label(self) -> None:
        if not hasattr(self, "subtotal_label"):
            return

        base_total = self._subtotal
        discount_percent = self._discount_percent if self.is_enabled() else 0.0
        markup_percent = self._markup_percent if self.is_enabled() else 0.0
        discount_amount = base_total * (discount_percent / 100.0)
        markup_amount = base_total * (markup_percent / 100.0)
        final_total = self.get_subtotal()

        parts: list[str] = []
        if discount_percent > 0:
            parts.append(
                f"− {self._format_currency(discount_amount)} ({self._format_percent(discount_percent)})"
            )
        if markup_percent > 0:
            parts.append(
                f"+ {self._format_currency(markup_amount)} ({self._format_percent(markup_percent)})"
            )

        prefix_amount = base_total if self.is_enabled() else self.get_subtotal()
        prefix = f"{tr('Промежуточная сумма', self.lang)}: {self._format_currency(prefix_amount)}"
        if parts:
            self.subtotal_label.setText(
                f"{prefix} {' '.join(parts)} = {self._format_currency(final_total)}"
            )
        else:
            self.subtotal_label.setText(prefix)

    def get_discount_percent(self) -> float:
        return self._discount_percent

    def set_discount_percent(self, value: float) -> None:
        self._discount_percent = max(0.0, min(100.0, float(value)))
        self._update_discount_label()
        self.subtotal_changed.emit(self.get_subtotal())

    def get_markup_percent(self) -> float:
        return self._markup_percent

    def set_markup_percent(self, value: float) -> None:
        self._markup_percent = max(0.0, min(100.0, float(value)))
        self._update_discount_label()
        self.subtotal_changed.emit(self.get_subtotal())

    def _format_currency(self, value: float) -> str:
        formatted = format_amount(value, self.lang)
        symbol = self.currency_symbol or ""
        if not symbol:
            return formatted
        if symbol.isalpha():
            return f"{formatted} {symbol}"
        return f"{formatted}{symbol}"

    @staticmethod
    def _format_percent(value: float) -> str:
        if abs(value - round(value)) < 1e-6:
            return f"{int(round(value))}%"
        return f"{value:.1f}%"

    def _show_modifiers_menu(self) -> None:
        if not self.is_enabled():
            return

        menu = QMenu(self)
        menu.setSeparatorsCollapsible(False)

        button_group = QButtonGroup(menu)
        button_group.setExclusive(True)

        def create_radio_action(text: str, with_spin: bool = False) -> tuple[QWidgetAction, QRadioButton, QSpinBox | None]:
            container = QWidget(menu)
            layout = QHBoxLayout(container)
            layout.setContentsMargins(*GROUP_SECTION_MARGINS)
            layout.setSpacing(GROUP_SECTION_SPACING)
            radio = QRadioButton(text)
            layout.addWidget(radio)
            button_group.addButton(radio)
            spin: QSpinBox | None = None
            if with_spin:
                spin = QSpinBox(container)
                spin.setRange(0, 100)
                spin.setSuffix("%")
                spin.setSingleStep(1)
                spin.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
                layout.addWidget(spin)
            layout.addStretch()
            action = QWidgetAction(menu)
            action.setDefaultWidget(container)
            menu.addAction(action)
            return action, radio, spin

        _, none_radio, _ = create_radio_action(tr("Без модификаторов", self.lang))
        _, discount_radio, discount_spin = create_radio_action(tr("Скидка", self.lang), True)
        _, markup_radio, markup_spin = create_radio_action(tr("Наценка", self.lang), True)

        menu.addSeparator()

        totals_widget = QWidget(menu)
        totals_layout = QHBoxLayout(totals_widget)
        totals_layout.setContentsMargins(*GROUP_SECTION_MARGINS)
        totals_layout.setSpacing(GROUP_SECTION_SPACING)
        totals_label = QLabel()
        totals_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        totals_layout.addWidget(totals_label)
        totals_action = QWidgetAction(menu)
        totals_action.setDefaultWidget(totals_widget)
        menu.addAction(totals_action)

        buttons_widget = QWidget(menu)
        buttons_layout = QHBoxLayout(buttons_widget)
        buttons_layout.setContentsMargins(*GROUP_SECTION_MARGINS)
        buttons_layout.setSpacing(GROUP_SECTION_SPACING)
        apply_button = QPushButton(tr("Применить", self.lang))
        cancel_button = QPushButton(tr("Отмена", self.lang))
        buttons_layout.addStretch()
        buttons_layout.addWidget(apply_button)
        buttons_layout.addWidget(cancel_button)
        buttons_action = QWidgetAction(menu)
        buttons_action.setDefaultWidget(buttons_widget)
        menu.addAction(buttons_action)

        base_total = self._subtotal

        if self._discount_percent > 0 and self._markup_percent <= 0:
            discount_radio.setChecked(True)
        elif self._markup_percent > 0 and self._discount_percent <= 0:
            markup_radio.setChecked(True)
        elif self._discount_percent == 0 and self._markup_percent == 0:
            none_radio.setChecked(True)
        else:
            discount_radio.setChecked(True)

        if discount_spin:
            discount_spin.setEnabled(discount_radio.isChecked())
            discount_spin.setValue(int(round(self._discount_percent)))
        if markup_spin:
            markup_spin.setEnabled(markup_radio.isChecked())
            markup_spin.setValue(int(round(self._markup_percent)))

        def update_totals_label():
            discount_value = float(discount_spin.value()) if discount_spin and discount_radio.isChecked() else 0.0
            markup_value = float(markup_spin.value()) if markup_spin and markup_radio.isChecked() else 0.0
            preview_total = base_total - base_total * (discount_value / 100.0) + base_total * (markup_value / 100.0)
            totals_label.setText(
                f"{tr('Итого', self.lang)}: {self._format_currency(preview_total)}"
            )

        update_totals_label()

        def on_button_toggled(button, checked):
            if button is discount_radio and discount_spin:
                discount_spin.setEnabled(checked)
            if button is markup_radio and markup_spin:
                markup_spin.setEnabled(checked)
            update_totals_label()

        button_group.buttonToggled.connect(on_button_toggled)  # type: ignore[arg-type]

        if discount_spin:
            discount_spin.valueChanged.connect(update_totals_label)
        if markup_spin:
            markup_spin.valueChanged.connect(update_totals_label)

        def apply_changes():
            if none_radio.isChecked():
                self._discount_percent = 0.0
                self._markup_percent = 0.0
            elif discount_radio.isChecked() and discount_spin:
                self._discount_percent = float(discount_spin.value())
                self._markup_percent = 0.0
            elif markup_radio.isChecked() and markup_spin:
                self._discount_percent = 0.0
                self._markup_percent = float(markup_spin.value())
            self._update_discount_label()
            self.subtotal_changed.emit(self.get_subtotal())
            menu.close()

        def cancel_changes():
            menu.close()

        apply_button.clicked.connect(apply_changes)
        cancel_button.clicked.connect(cancel_changes)

        button = getattr(self, "modifiers_button", None)
        if button:
            menu.exec(button.mapToGlobal(button.rect().bottomLeft()))
        else:
            menu.exec(self.mapToGlobal(self.rect().center()))

    # --------------------------------------------------------------- data i/o
    def get_data(self) -> Dict:
        if not self.is_enabled():
            return {}
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
        self._update_discount_label()

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
        self.set_currency(self.currency_symbol, self.currency_code)
        self.update_sums()
        self._update_discount_label()

    def is_enabled(self) -> bool:
        return getattr(self, "enabled_checkbox", None) is None or self.enabled_checkbox.isChecked()

    def _on_enabled_toggled(self, checked: bool) -> None:
        self.header_edit.setEnabled(checked)
        self.table.setEnabled(checked)
        self.subtotal_label.setEnabled(checked)
        if hasattr(self, "modifiers_button"):
            self.modifiers_button.setEnabled(checked)
        self._update_discount_label()
        self.subtotal_changed.emit(self.get_subtotal())


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

