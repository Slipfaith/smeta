from typing import List, Dict, Any

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QGroupBox, QTableWidget, QTableWidgetItem,
    QLabel, QHeaderView, QSizePolicy, QMenu, QHBoxLayout
)
from PySide6.QtCore import Qt, Signal
from .utils import format_rate, _to_float
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
        self._setup_ui(initial_volume)

    def _setup_ui(self, initial_volume: float):
        layout = QVBoxLayout()

        header = QHBoxLayout()
        self.title_label = QLabel()
        header.addWidget(self.title_label)
        header.addStretch()
        layout.addLayout(header)

        group = QGroupBox()
        group.setCheckable(False)
        vbox = QVBoxLayout()

        self.table = QTableWidget(1, 4)
        self.table.setHorizontalHeaderLabels([
            tr("Названия работ", self.lang),
            tr("Час", self.lang),
            f"{tr('Ставка', self.lang)} ({self.currency_symbol})",
            f"{tr('Сумма', self.lang)} ({self.currency_symbol})",
        ])
        self.table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.table.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self.table.setWordWrap(False)

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
            if self.rows_deleted[row]:
                del_act.setEnabled(False)
            else:
                restore_act.setEnabled(False)
            if sum(1 for d in self.rows_deleted if not d) <= 1:
                del_act.setEnabled(False)
            action = menu.exec(self.table.mapToGlobal(pos))
            if action == add_act:
                self.add_row_after(row)
            elif action == del_act:
                self.remove_row_at(row)
            elif action == restore_act:
                self.restore_row_at(row)

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(show_menu)

        self.subtotal_label = QLabel(f"{tr('Промежуточная сумма', self.lang)}: 0.00 {self.currency_symbol}")
        self.subtotal_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        vbox.addWidget(self.subtotal_label)

        group.setLayout(vbox)
        layout.addWidget(group)
        self.setLayout(layout)

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
                total_item.setText(f"{total:.2f}")
                self.table.blockSignals(False)
                subtotal += total
            self.subtotal_label.setText(f"{tr('Промежуточная сумма', self.lang)}: {subtotal:.2f} {self.currency_symbol}")
            self._subtotal = subtotal
            self.subtotal_changed.emit(subtotal)
            self._fit_table_height(self.table)
        except Exception:
            pass

    # ---------- accessors ----------
    def get_subtotal(self) -> float:
        return self._subtotal

    def get_data(self) -> List[Dict[str, Any]]:
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

    def load_data(self, rows: List[Dict[str, Any]]):
        self.table.blockSignals(True)
        self.table.setRowCount(len(rows))
        self.rows_deleted = [False] * len(rows)
        for i, row_data in enumerate(rows):
            self.table.setItem(i, 0, QTableWidgetItem(row_data.get("parameter", "")))
            self.table.setItem(i, 1, QTableWidgetItem(str(row_data.get("volume", 0))))
            sep = "." if self.lang == "en" else None
            self.table.setItem(i, 2, QTableWidgetItem(format_rate(row_data.get('rate', 0), sep)))
            total_item = QTableWidgetItem(f"{row_data.get('total', 0):.2f}")
            total_item.setFlags(Qt.ItemIsEnabled)
            self.table.setItem(i, 3, total_item)
            self._set_row_deleted(i, False)
        self.table.blockSignals(False)
        self._fit_table_height(self.table)
        self.update_sums()

    def set_currency(self, symbol: str, code: str):
        self.currency_symbol = symbol
        self.currency_code = code
        self.table.setHorizontalHeaderLabels([
            tr("Названия работ", self.lang),
            tr("Ед-ца", self.lang),
            f"{tr('Ставка', self.lang)} ({symbol})",
            f"{tr('Сумма', self.lang)} ({symbol})",
        ])
        self.update_sums()

    def set_language(self, lang: str):
        self.lang = lang
        self.title_label.setText(tr("Запуск и управление проектом", lang))
        self.set_currency(self.currency_symbol, self.currency_code)
        item = self.table.item(0, 0)
        if item:
            item.setText(tr("Запуск и управление проектом", lang))
        self.update_sums()
