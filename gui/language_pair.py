from typing import Dict, List, Any

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QGroupBox, QTableWidget, QTableWidgetItem, QLabel,
    QHeaderView, QSizePolicy, QHBoxLayout, QPushButton
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont

from logic.service_config import ServiceConfig


class LanguagePairWidget(QWidget):
    """Виджет для одной языковой пары (только Перевод)"""

    remove_requested = Signal()

    def __init__(self, pair_name: str):
        super().__init__()
        self.pair_name = pair_name
        self.setup_ui()

    # ---------------- UI ----------------
    def setup_ui(self):
        layout = QVBoxLayout()

        header = QHBoxLayout()
        title = QLabel(f"Языковая пара: {self.pair_name}")
        title.setFont(QFont("Arial", 10, QFont.Bold))
        header.addWidget(title)
        header.addStretch()
        remove_btn = QPushButton("Удалить")
        remove_btn.clicked.connect(self.remove_requested.emit)
        header.addWidget(remove_btn)
        layout.addLayout(header)

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

        vbox = QVBoxLayout()

        table = QTableWidget(len(rows), 4)
        table.setHorizontalHeaderLabels(["Параметр", "Объем", "Ставка (руб)", "Сумма (руб)"])

        # ВАЖНО: никаких локальных скроллов — всё видно сразу
        table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        table.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        table.setWordWrap(False)

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

        # Автоподгон ширин
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)

        # Пересчёт ставок/сумм
        table.itemChanged.connect(lambda item: self.update_rates_and_sums(table, rows, base_rate_row))

        vbox.addWidget(table)

        # кнопки добавления/удаления строк
        controls = QHBoxLayout()
        add_btn = QPushButton("Добавить строку")
        del_btn = QPushButton("Удалить строку")
        controls.addWidget(add_btn)
        controls.addWidget(del_btn)
        controls.addStretch()
        vbox.addLayout(controls)

        def add_row():
            r = table.rowCount()
            table.insertRow(r)
            table.setItem(r, 0, QTableWidgetItem("Новая строка"))
            table.setItem(r, 1, QTableWidgetItem("0"))
            table.setItem(r, 2, QTableWidgetItem("0.00"))
            sum_item = QTableWidgetItem("0.00")
            sum_item.setFlags(Qt.ItemIsEnabled)
            table.setItem(r, 3, sum_item)
            rows.append({"name": "Новая строка", "is_base": False, "multiplier": 1.0})
            self.update_rates_and_sums(table, rows, base_rate_row)
            self._fit_table_height(table)

        def del_row():
            nonlocal base_rate_row
            row = table.currentRow()
            if row >= 0:
                table.removeRow(row)
                if 0 <= row < len(rows):
                    rows.pop(row)
                if base_rate_row is not None:
                    if row == base_rate_row:
                        base_rate_row = None
                    elif row < base_rate_row:
                        base_rate_row -= 1
                setattr(group, 'base_rate_row', base_rate_row)
                self.update_rates_and_sums(table, rows, base_rate_row)
                self._fit_table_height(table)

        add_btn.clicked.connect(add_row)
        del_btn.clicked.connect(del_row)

        # Промежуточная сумма
        subtotal_label = QLabel("Промежуточная сумма: 0.00 ₽")
        subtotal_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        subtotal_label.setObjectName("subtotal_label")
        vbox.addWidget(subtotal_label)

        group.setLayout(vbox)

        # сохранить ссылки
        setattr(group, 'table', table)
        setattr(group, 'rows_config', rows)
        setattr(group, 'base_rate_row', base_rate_row)
        setattr(group, 'subtotal_label', subtotal_label)

        # начальный пересчёт + авто-раскрытие высоты
        self.update_rates_and_sums(table, rows, base_rate_row)
        self._fit_table_height(table)

        return group

    # ---------------- Logic ----------------
    def _fit_table_height(self, table: QTableWidget):
        """Делает таблицу фиксированной высоты по всем строкам (без внутреннего скролла)."""
        header_h = table.horizontalHeader().height()
        rows_h = sum(table.rowHeight(r) for r in range(table.rowCount()))
        frame = table.frameWidth() * 2
        # небольшой запас под сетку/паддинги
        total = header_h + rows_h + frame + 2
        table.setMinimumHeight(total)
        table.setMaximumHeight(total)

    def update_rates_and_sums(self, table: QTableWidget, rows: List[Dict], base_rate_row: int):
        try:
            base_rate = 0.0
            if base_rate_row is not None and table.item(base_rate_row, 2):
                base_rate = float(table.item(base_rate_row, 2).text() or "0")

            subtotal = 0.0
            for row in range(table.rowCount()):
                row_cfg = rows[row]

                # авто-ставки для непервой строки
                if not row_cfg["is_base"] and base_rate_row is not None:
                    auto_rate = base_rate * row_cfg["multiplier"]
                    if table.item(row, 2):
                        table.blockSignals(True)
                        table.item(row, 2).setText(f"{auto_rate:.2f}")
                        table.blockSignals(False)

                volume = float((table.item(row, 1).text() if table.item(row, 1) else "0") or "0")
                rate = float((table.item(row, 2).text() if table.item(row, 2) else "0") or "0")
                total = volume * rate
                if table.item(row, 3):
                    table.blockSignals(True)
                    table.item(row, 3).setText(f"{total:.2f}")
                    table.blockSignals(False)
                subtotal += total

            # обновить «Промежуточную сумму»
            parent_group: QGroupBox = self.translation_group
            lbl: QLabel = getattr(parent_group, 'subtotal_label', None)
            if lbl:
                lbl.setText(f"Промежуточная сумма: {subtotal:.2f} ₽")

            # после любых изменений гарантируем отсутствие локального скролла
            self._fit_table_height(table)

        except (ValueError, AttributeError):
            pass

    def refresh_layout(self):
        """Публичный хук для внешнего кода: раскрыть таблицы по высоте ещё раз."""
        if hasattr(self.translation_group, "table"):
            self._fit_table_height(self.translation_group.table)

    # ---------------- Data ----------------
    def get_data(self) -> Dict[str, Any]:
        data = {"pair_name": self.pair_name, "services": {}}
        if self.translation_group.isChecked():
            data["services"]["translation"] = self._get_table_data(self.translation_group.table)
        return data

    def _get_table_data(self, table: QTableWidget) -> List[Dict[str, Any]]:
        out = []
        for row in range(table.rowCount()):
            out.append({
                "parameter": table.item(row, 0).text() if table.item(row, 0) else "",
                "volume": float((table.item(row, 1).text() if table.item(row, 1) else "0") or "0"),
                "rate":   float((table.item(row, 2).text() if table.item(row, 2) else "0") or "0"),
                "total":  float((table.item(row, 3).text() if table.item(row, 3) else "0") or "0"),
            })
        return out

    def load_table_data(self, data: List[Dict[str, Any]]):
        group = self.translation_group
        table = group.table
        rows = group.rows_config
        base_rate_row = group.base_rate_row

        if len(data) > table.rowCount():
            for _ in range(len(data) - table.rowCount()):
                r = table.rowCount()
                table.insertRow(r)
                table.setItem(r, 0, QTableWidgetItem("Новая строка"))
                table.setItem(r, 1, QTableWidgetItem("0"))
                table.setItem(r, 2, QTableWidgetItem("0.00"))
                sum_item = QTableWidgetItem("0.00")
                sum_item.setFlags(Qt.ItemIsEnabled)
                table.setItem(r, 3, sum_item)
                rows.append({"name": "Новая строка", "is_base": False, "multiplier": 1.0})

        for row, row_data in enumerate(data):
            if row < table.rowCount():
                table.item(row, 0).setText(row_data.get("parameter", ""))
                table.item(row, 1).setText(str(row_data.get("volume", 0)))
                table.item(row, 2).setText(str(row_data.get("rate", 0)))
                table.item(row, 3).setText(str(row_data.get("total", 0)))

        self.update_rates_and_sums(table, rows, base_rate_row)
        self._fit_table_height(table)
