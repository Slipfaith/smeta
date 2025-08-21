from typing import Dict, List, Any

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QGroupBox, QTableWidget, QTableWidgetItem, QLabel,
    QHeaderView, QSizePolicy, QHBoxLayout, QPushButton, QMenu
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

        rows = [dict(r, deleted=False) for r in rows]

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
                rate_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            else:
                if base_rate_row is None:
                    base_rate_row = i
            table.setItem(i, 2, rate_item)

            sum_item = QTableWidgetItem("0.00")
            sum_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            table.setItem(i, 3, sum_item)

        # Автоподгон ширин
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)

        # Пересчёт ставок/сумм
        table.itemChanged.connect(
            lambda item, t=table, r=rows, g=group: self.update_rates_and_sums(t, r, getattr(g, 'base_rate_row'))
        )

        vbox.addWidget(table)

        # контекстное меню для добавления/удаления/восстановления строк
        def show_menu(pos):
            row = table.rowAt(pos.y())
            if row < 0:
                return
            menu = QMenu(table)
            add_act = menu.addAction("Добавить строку")
            del_act = menu.addAction("Удалить строку")
            restore_act = menu.addAction("Восстановить строку")
            row_cfg = rows[row]
            if row_cfg.get("deleted"):
                del_act.setEnabled(False)
            else:
                restore_act.setEnabled(False)
            action = menu.exec(table.mapToGlobal(pos))
            if action == add_act:
                self._add_row_after(table, rows, group, row)
            elif action == del_act:
                self._delete_row(table, rows, group, row)
            elif action == restore_act:
                self._restore_row(table, rows, group, row)

        table.setContextMenuPolicy(Qt.CustomContextMenu)
        table.customContextMenuRequested.connect(show_menu)

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

    def _add_row_after(self, table: QTableWidget, rows: List[Dict], group: QGroupBox, row: int):
        base_rate_row = getattr(group, 'base_rate_row', None)
        insert_at = row + 1
        table.insertRow(insert_at)
        table.setItem(insert_at, 0, QTableWidgetItem("Новая строка"))
        table.setItem(insert_at, 1, QTableWidgetItem("0"))
        rate_item = QTableWidgetItem("0.00")
        rate_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        table.setItem(insert_at, 2, rate_item)
        sum_item = QTableWidgetItem("0.00")
        sum_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        table.setItem(insert_at, 3, sum_item)
        rows.insert(insert_at, {"name": "Новая строка", "is_base": False, "multiplier": 1.0, "deleted": False})
        if base_rate_row is not None and insert_at <= base_rate_row:
            base_rate_row += 1
            setattr(group, 'base_rate_row', base_rate_row)
        self.update_rates_and_sums(table, rows, base_rate_row)
        self._fit_table_height(table)

    def _set_row_deleted(self, table: QTableWidget, rows: List[Dict], row: int, deleted: bool):
        rows[row]['deleted'] = deleted
        for col in range(table.columnCount()):
            item = table.item(row, col)
            if not item:
                continue
            if deleted:
                item.setForeground(Qt.gray)
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            else:
                item.setForeground(Qt.black)
                if col == 2:
                    if rows[row]['is_base']:
                        item.setFlags(Qt.ItemIsEditable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                    else:
                        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                elif col == 3:
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                else:
                    item.setFlags(Qt.ItemIsEditable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)

    def _delete_row(self, table: QTableWidget, rows: List[Dict], group: QGroupBox, row: int):
        base_rate_row = getattr(group, 'base_rate_row', None)
        self._set_row_deleted(table, rows, row, True)
        if base_rate_row == row:
            base_rate_row = None
            setattr(group, 'base_rate_row', base_rate_row)
        self.update_rates_and_sums(table, rows, base_rate_row)

    def _restore_row(self, table: QTableWidget, rows: List[Dict], group: QGroupBox, row: int):
        base_rate_row = getattr(group, 'base_rate_row', None)
        self._set_row_deleted(table, rows, row, False)
        if rows[row]['is_base'] and base_rate_row is None:
            base_rate_row = row
            setattr(group, 'base_rate_row', base_rate_row)
        self.update_rates_and_sums(table, rows, base_rate_row)

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
            if base_rate_row is not None and rows[base_rate_row].get('deleted'):
                base_rate_row = None
            if base_rate_row is not None and table.item(base_rate_row, 2):
                base_rate = float(table.item(base_rate_row, 2).text() or "0")

            subtotal = 0.0
            for row in range(table.rowCount()):
                row_cfg = rows[row]
                if row_cfg.get('deleted'):
                    if table.item(row, 3):
                        table.blockSignals(True)
                        table.item(row, 3).setText("0.00")
                        table.blockSignals(False)
                    continue

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
        rows_cfg = self.translation_group.rows_config
        for row in range(table.rowCount()):
            if rows_cfg[row].get('deleted'):
                continue
            out.append({
                "parameter": table.item(row, 0).text() if table.item(row, 0) else "",
                "volume": float((table.item(row, 1).text() if table.item(row, 1) else "0") or "0"),
                "rate":   float((table.item(row, 2).text() if table.item(row, 2) else "0") or "0"),
                "total":  float((table.item(row, 3).text() if table.item(row, 3) else "0") or "0"),
                "is_base": rows_cfg[row].get("is_base", False),
                "multiplier": rows_cfg[row].get("multiplier"),
            })
        return out

    def load_table_data(self, data: List[Dict[str, Any]]):
        group = self.translation_group
        table = group.table
        rows = group.rows_config
        base_rate_row = None

        if len(data) > table.rowCount():
            for _ in range(len(data) - table.rowCount()):
                r = table.rowCount()
                table.insertRow(r)
                table.setItem(r, 0, QTableWidgetItem("Новая строка"))
                table.setItem(r, 1, QTableWidgetItem("0"))
                rate_item = QTableWidgetItem("0.00")
                rate_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                table.setItem(r, 2, rate_item)
                sum_item = QTableWidgetItem("0.00")
                sum_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                table.setItem(r, 3, sum_item)
                rows.append({"name": "Новая строка", "is_base": False, "multiplier": 1.0, "deleted": False})

        for row, row_data in enumerate(data):
            if row < table.rowCount():
                table.item(row, 0).setText(row_data.get("parameter", ""))
                table.item(row, 1).setText(str(row_data.get("volume", 0)))
                table.item(row, 2).setText(str(row_data.get("rate", 0)))
                table.item(row, 3).setText(str(row_data.get("total", 0)))
                rows[row]["is_base"] = row_data.get("is_base", rows[row].get("is_base", False))
                rows[row]["multiplier"] = row_data.get("multiplier", rows[row].get("multiplier", 1.0))
                if rows[row].get("is_base"):
                    base_rate_row = row

        setattr(group, 'base_rate_row', base_rate_row)
        self.update_rates_and_sums(table, rows, base_rate_row)
        self._fit_table_height(table)
