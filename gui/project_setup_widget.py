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
    QDoubleSpinBox,
)
from PySide6.QtCore import Qt, Signal
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
            add_act = menu.addAction(tr("Добавить строку", self.lang))
            del_act = menu.addAction(tr("Удалить строку", self.lang))
            restore_act = menu.addAction(tr("Восстановить строку", self.lang))
            del_selected_act = menu.addAction(
                tr("Удалить выбранные строки", self.lang)
            )
            if self.rows_deleted[row]:
                del_act.setEnabled(False)
            else:
                restore_act.setEnabled(False)
            if sum(1 for d in self.rows_deleted if not d) <= 1:
                del_act.setEnabled(False)
            selectable = [
                idx
                for idx in {index.row() for index in self.table.selectedIndexes()}
                if 0 <= idx < len(self.rows_deleted) and not self.rows_deleted[idx]
            ]
            if len(selectable) <= 1 or sum(
                1 for d in self.rows_deleted if not d
            ) - len(selectable) < 1:
                del_selected_act.setEnabled(False)
            action = menu.exec(self.table.mapToGlobal(pos))
            if action == add_act:
                self.add_row_after(row)
            elif action == del_act:
                self.remove_row_at(row)
            elif action == restore_act:
                self.restore_row_at(row)
            elif action == del_selected_act:
                self.remove_selected_rows()

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(show_menu)

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
            f"{tr('Сумма скидки', self.lang)}: 0.00 {self.currency_symbol}"
        )
        self.discounted_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        discount_layout.addWidget(self.discounted_label)
        vbox.addLayout(discount_layout)

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
            f"{tr('Сумма наценки', self.lang)}: 0.00 {self.currency_symbol}"
        )
        self.markup_amount_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        markup_layout.addWidget(self.markup_amount_label)
        vbox.addLayout(markup_layout)

        self.subtotal_label = QLabel(
            f"{tr('Промежуточная сумма', self.lang)}: 0.00 {self.currency_symbol}"
        )
        self.subtotal_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        vbox.addWidget(self.subtotal_label)

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
        if row >= 0 and not self.rows_deleted[row]:
            if sum(1 for d in self.rows_deleted if not d) <= 1:
                return
            self._set_row_deleted(row, True)
            self.update_sums()

    def remove_selected_rows(self):
        active_indices = [i for i, deleted in enumerate(self.rows_deleted) if not deleted]
        selected = sorted({index.row() for index in self.table.selectedIndexes()})
        removable = [
            row
            for row in selected
            if 0 <= row < len(self.rows_deleted) and not self.rows_deleted[row]
        ]
        if len(removable) <= 1:
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
        if row >= 0 and self.rows_deleted[row]:
            self._set_row_deleted(row, False)
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
                    item.setText("0.00")
            else:
                item.setForeground(Qt.black)
                flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable
                if col != 3:
                    flags |= Qt.ItemIsEditable
                item.setFlags(flags)

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
        if hasattr(self, "discount_spin"):
            self.discount_spin.setEnabled(checked)
        if hasattr(self, "discount_label"):
            self.discount_label.setEnabled(checked)
        if hasattr(self, "discounted_label"):
            self.discounted_label.setEnabled(checked)
        if hasattr(self, "markup_spin"):
            self.markup_spin.setEnabled(checked)
        if hasattr(self, "markup_label"):
            self.markup_label.setEnabled(checked)
        if hasattr(self, "markup_amount_label"):
            self.markup_amount_label.setEnabled(checked)
        self.update_sums()

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
        if hasattr(self, "discount_label"):
            self.discount_label.setText(tr("Скидка, %", self.lang))
        if hasattr(self, "markup_label"):
            self.markup_label.setText(tr("Наценка, %", self.lang))
        suffix = f" {self.currency_symbol}" if self.currency_symbol else ""
        discount_amount = (
            self._subtotal * (self._discount_percent / 100.0)
            if self.is_enabled()
            else 0.0
        )
        markup_amount = (
            self._subtotal * (self._markup_percent / 100.0)
            if self.is_enabled()
            else 0.0
        )
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
        self.update_sums()
