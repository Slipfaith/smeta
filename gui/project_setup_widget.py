from typing import List, Dict, Any

from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QGroupBox,
    QTableWidget,
    QTableWidgetItem,
    QLabel,
    QHeaderView,
    QSizePolicy,
    QMenu,
    QHBoxLayout,
    QAbstractItemView,
    QPushButton,
    QSpinBox,
    QRadioButton,
    QWidgetAction,
    QButtonGroup,
)
from PySide6.QtCore import Qt, Signal
from gui.styles import GROUP_SECTION_MARGINS, GROUP_SECTION_SPACING
from .utils import format_rate, _to_float, format_amount
from logic.translation_config import tr


class ProjectSetupWidget(QWidget):
    """Table for project setup and management costs."""

    remove_requested = Signal()
    subtotal_changed = Signal(float)

    def __init__(self, initial_volume: float = 0.0, currency_symbol: str = "₽", currency_code: str = "RUB", lang: str = "ru"):
        super().__init__()
        self.currency_symbol = currency_symbol
        self.currency_code = currency_code
        self.lang = lang
        self._subtotal = 0.0
        self._discount_percent = 0.0
        self._markup_percent = 0.0
        self._setup_ui(initial_volume)

    def _setup_ui(self, initial_volume: float):
        layout = QVBoxLayout()

        header = QHBoxLayout()
        self.title_label = QLabel()
        header.addWidget(self.title_label)
        header.addStretch()
        layout.addLayout(header)

        group = QGroupBox()
        group.setCheckable(True)
        group.setChecked(True)
        vbox = QVBoxLayout()

        self.table = QTableWidget(1, 4)
        self.table.setHorizontalHeaderLabels([
            tr("Названия работ", self.lang),
            tr("Объем", self.lang),
            f"{tr('Ставка', self.lang)} ({self.currency_symbol})",
            f"{tr('Сумма', self.lang)} ({self.currency_symbol})",
        ])
        self.table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.table.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self.table.setWordWrap(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)

        self.table.setItem(0, 0, QTableWidgetItem(tr("Запуск и управление проектом", self.lang)))
        self.table.setItem(0, 1, QTableWidgetItem(str(initial_volume)))
        self.table.setItem(0, 2, QTableWidgetItem("0.00"))
        total_item = QTableWidgetItem("0.00")
        total_item.setFlags(Qt.ItemIsEnabled)
        self.table.setItem(0, 3, total_item)
        self.rows_deleted = [False]
        self._set_row_deleted(0, False)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)

        self.table.itemChanged.connect(self.update_sums)

        vbox.addWidget(self.table)

        def show_menu(pos):
            row = self.table.rowAt(pos.y())
            if row < 0:
                row = self.table.rowCount() - 1
            menu = QMenu(self.table)
            add_act = menu.addAction(tr("Добавить строку", self.lang))
            del_act = menu.addAction(tr("Удалить", self.lang))
            restore_act = menu.addAction(tr("Восстановить", self.lang))

            active_rows = sum(1 for d in self.rows_deleted if not d)
            selectable = sorted(
                {
                    index.row()
                    for index in self.table.selectedIndexes()
                    if 0 <= index.row() < len(self.rows_deleted)
                    and not self.rows_deleted[index.row()]
                }
            )
            selected_deleted = sorted(
                {
                    index.row()
                    for index in self.table.selectedIndexes()
                    if 0 <= index.row() < len(self.rows_deleted)
                    and self.rows_deleted[index.row()]
                }
            )

            if not (0 <= row < len(self.rows_deleted)):
                del_act.setEnabled(False)
                restore_act.setEnabled(False)

            can_delete = False
            if selectable:
                can_delete = active_rows - len(selectable) >= 1
            elif 0 <= row < len(self.rows_deleted):
                can_delete = (not self.rows_deleted[row]) and active_rows > 1
            if not can_delete:
                del_act.setEnabled(False)

            can_restore = False
            if selected_deleted:
                can_restore = True
            elif 0 <= row < len(self.rows_deleted) and self.rows_deleted[row]:
                can_restore = True
            if not can_restore:
                restore_act.setEnabled(False)
            action = menu.exec(self.table.mapToGlobal(pos))
            if action == add_act:
                self.add_row_after(row)
            elif action == del_act:
                targets = selectable if selectable else [row]
                self.remove_rows(targets)
            elif action == restore_act:
                targets = selected_deleted if selected_deleted else [row]
                self.restore_rows(targets)

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(show_menu)

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

        vbox.addLayout(subtotal_layout)

        group.setLayout(vbox)
        layout.addWidget(group)
        self.setLayout(layout)
        group.toggled.connect(self._on_group_toggled)

        self.group_box = group
        self._fit_table_height(self.table)
        self.update_sums()
        self.set_language(self.lang)

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
        self.table.setItem(insert_at, 0, QTableWidgetItem(tr("Новая строка", self.lang)))
        self.table.setItem(insert_at, 1, QTableWidgetItem("0"))
        self.table.setItem(insert_at, 2, QTableWidgetItem("0.00"))
        total_item = QTableWidgetItem("0.00")
        total_item.setFlags(Qt.ItemIsEnabled)
        self.table.setItem(insert_at, 3, total_item)
        self.rows_deleted.insert(insert_at, False)
        self._set_row_deleted(insert_at, False)
        self._fit_table_height(self.table)
        self.update_sums()

    def remove_row(self):
        self.remove_row_at(self.table.currentRow())

    def remove_row_at(self, row: int):
        self.remove_rows([row])

    def remove_rows(self, rows: list[int]):
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
        if max_remove <= 0:
            return
        if len(removable) > max_remove:
            removable = removable[-max_remove:]
        for row in removable:
            self._set_row_deleted(row, True)
        self.update_sums()

    def restore_row_at(self, row: int):
        self.restore_rows([row])

    def restore_rows(self, rows: list[int]):
        restored = False
        for row in sorted({r for r in rows if 0 <= r < len(self.rows_deleted) and self.rows_deleted[r]}):
            self._set_row_deleted(row, False)
            restored = True
        if restored:
            self.update_sums()

    def _set_row_deleted(self, row: int, deleted: bool):
        self.rows_deleted[row] = deleted
        for col in range(self.table.columnCount()):
            item = self.table.item(row, col)
            if not item:
                continue
            if deleted:
                item.setForeground(Qt.gray)
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                if col == 3:
                    item.setText(format_amount(0.0, self.lang))
            else:
                item.setForeground(Qt.black)
                flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable
                if col != 3:
                    flags |= Qt.ItemIsEditable
                item.setFlags(flags)
            font = item.font()
            font.setStrikeOut(deleted)
            item.setFont(font)

    def set_volume(self, value: float):
        if self.table.rowCount() == 0:
            self.add_row()
        item = self.table.item(0, 1)
        self.table.blockSignals(True)
        item.setText(f"{value}")
        self.table.blockSignals(False)
        self.update_sums()

    def is_enabled(self) -> bool:
        return self.group_box.isChecked()

    def _on_group_toggled(self, checked: bool):
        self.table.setEnabled(checked)
        self.subtotal_label.setEnabled(checked)
        self.title_label.setEnabled(checked)
        if hasattr(self, "modifiers_button"):
            self.modifiers_button.setEnabled(checked)
        self.update_sums()
        self._update_discount_label()

    # ---------- calculations ----------
    def update_sums(self):
        try:
            subtotal = 0.0
            for row in range(self.table.rowCount()):
                if self.rows_deleted[row]:
                    continue
                volume_item = self.table.item(row, 1)
                rate_item = self.table.item(row, 2)
                volume = _to_float(volume_item.text() if volume_item else "0")
                rate_text = rate_item.text() if rate_item else "0"
                if self.lang == "en":
                    sep = "."
                else:
                    sep = "," if "," in rate_text else "."
                rate = _to_float(rate_text)
                self.table.blockSignals(True)
                rate_item.setText(format_rate(rate_text, sep))
                self.table.blockSignals(False)
                total = volume * rate
                total_item = self.table.item(row, 3)
                if total_item is None:
                    total_item = QTableWidgetItem("0.00")
                    total_item.setFlags(Qt.ItemIsEnabled)
                    self.table.setItem(row, 3, total_item)
                self.table.blockSignals(True)
                total_item.setText(format_amount(total, self.lang))
                self.table.blockSignals(False)
                subtotal += total
            self._subtotal = subtotal
            self._update_discount_label()
            self.subtotal_changed.emit(self.get_subtotal())
            self._fit_table_height(self.table)
        except Exception:
            pass

    # ---------- accessors ----------
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

        prefix_amount = base_total if self.is_enabled() else final_total
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

    def get_data(self) -> List[Dict[str, Any]]:
        if not self.is_enabled():
            return []
        data: List[Dict[str, Any]] = []
        for row in range(self.table.rowCount()):
            if self.rows_deleted[row]:
                continue
            data.append({
                "parameter": self.table.item(row, 0).text() if self.table.item(row, 0) else "",
                "volume": _to_float(self.table.item(row, 1).text() if self.table.item(row, 1) else "0"),
                "rate": _to_float(self.table.item(row, 2).text() if self.table.item(row, 2) else "0"),
                "total": _to_float(self.table.item(row, 3).text() if self.table.item(row, 3) else "0"),
            })
        return data

    def load_data(self, rows: List[Dict[str, Any]], enabled: bool | None = None):
        if enabled is not None:
            self.group_box.setChecked(enabled)
        self.table.blockSignals(True)
        self.table.setRowCount(len(rows))
        self.rows_deleted = [False] * len(rows)
        for i, row_data in enumerate(rows):
            self.table.setItem(i, 0, QTableWidgetItem(row_data.get("parameter", "")))
            self.table.setItem(i, 1, QTableWidgetItem(str(row_data.get("volume", 0))))
            sep = "." if self.lang == "en" else None
            self.table.setItem(i, 2, QTableWidgetItem(format_rate(row_data.get('rate', 0), sep)))
            total_item = QTableWidgetItem(
                format_amount(row_data.get('total', 0), self.lang)
            )
            total_item.setFlags(Qt.ItemIsEnabled)
            self.table.setItem(i, 3, total_item)
            self._set_row_deleted(i, False)
        self.table.blockSignals(False)
        self._fit_table_height(self.table)
        self.update_sums()
        if enabled is not None:
            self._on_group_toggled(self.group_box.isChecked())

    def set_currency(self, symbol: str, code: str):
        self.currency_symbol = symbol
        self.currency_code = code
        rate_suffix = f" ({symbol})" if symbol else ""
        self.table.setHorizontalHeaderLabels([
            tr("Названия работ", self.lang),
            tr("Объем", self.lang),
            f"{tr('Ставка', self.lang)}{rate_suffix}",
            f"{tr('Сумма', self.lang)}{rate_suffix}",
        ])
        self.update_sums()
        self._update_discount_label()

    def convert_rates(self, multiplier: float):
        """Multiply all rate values by *multiplier* and update totals."""
        for row in range(self.table.rowCount()):
            if self.rows_deleted[row]:
                continue
            item = self.table.item(row, 2)
            if item is None:
                continue
            rate = _to_float(item.text())
            sep = '.' if self.lang == 'en' else ','
            item.setText(format_rate(rate * multiplier, sep))
        self.update_sums()

    def set_language(self, lang: str):
        self.lang = lang
        self.title_label.setText(tr("Запуск и управление проектом", lang))
        self.set_currency(self.currency_symbol, self.currency_code)
        item = self.table.item(0, 0)
        if item:
            item.setText(tr("Запуск и управление проектом", lang))
        self.update_sums()
        self._update_discount_label()
