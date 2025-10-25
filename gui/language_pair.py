from typing import Any, Dict, Iterable, List, Union

from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QGroupBox,
    QTableWidget,
    QTableWidgetItem,
    QLabel,
    QLineEdit,
    QHeaderView,
    QSizePolicy,
    QHBoxLayout,
    QMenu,
    QAbstractItemView,
    QPushButton,
    QSpinBox,
    QRadioButton,
    QButtonGroup,
    QWidgetAction,
    QToolButton,
    QStyle,
)
from PySide6.QtCore import Qt, Signal, QSize
from PySide6.QtGui import QFont

import copy

from logic.service_config import ServiceConfig
from logic.translation_config import tr
from gui.styles import (
    GROUP_SECTION_MARGINS,
    GROUP_SECTION_SPACING,
    LANGUAGE_PAIR_DELETE_BUTTON_STYLE,
    LANGUAGE_PAIR_DELETE_ICON_SIZE,
    REPORTS_LABEL_STYLE,
)
from .utils import format_rate, _to_float, format_amount


class LanguagePairWidget(QWidget):
    """Виджет для одной языковой пары (только Перевод)"""

    remove_requested = Signal()
    subtotal_changed = Signal(float)
    name_changed = Signal(str)

    def __init__(self, pair_name: str, currency_symbol: str = "₽", currency_code: str = "RUB", lang: str = "ru"):
        super().__init__()
        self.pair_name = pair_name
        self.currency_symbol = currency_symbol
        self.currency_code = currency_code
        self.lang = lang
        self.only_new_repeats_mode = False
        self._report_sources: List[str] = []
        # резерв для восстановления исходных значений объёмов/ставок
        self._backup_volumes = []
        self._backup_rates = []
        self._subtotal = 0.0
        self._discount_percent = 0.0
        self._markup_percent = 0.0
        self.setup_ui()

    # ---------------- UI ----------------
    def setup_ui(self):
        layout = QVBoxLayout()

        header = QHBoxLayout()
        self.title_label = QLabel()
        self.title_label.setFont(QFont("Arial", 10, QFont.Bold))
        self.title_edit = QLineEdit()
        self.title_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.title_edit.editingFinished.connect(self._on_title_edit)
        header.addWidget(self.title_label)
        header.addWidget(self.title_edit, 1)

        self.delete_button = QToolButton()
        self.delete_button.setIcon(self.style().standardIcon(QStyle.SP_TrashIcon))
        self.delete_button.setIconSize(QSize(*LANGUAGE_PAIR_DELETE_ICON_SIZE))
        self.delete_button.setAutoRaise(True)
        self.delete_button.setCursor(Qt.PointingHandCursor)
        self.delete_button.clicked.connect(self._on_delete_clicked)
        self.delete_button.setStyleSheet(LANGUAGE_PAIR_DELETE_BUTTON_STYLE)
        header.addWidget(self.delete_button)
        layout.addLayout(header)

        self.reports_label = QLabel()
        self.reports_label.setWordWrap(True)
        self.reports_label.setStyleSheet(REPORTS_LABEL_STYLE)
        self.reports_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.reports_label.hide()
        layout.addWidget(self.reports_label)

        self.services_layout = QVBoxLayout()

        # Только Перевод
        self.translation_group = self.create_service_group("Перевод", ServiceConfig.TRANSLATION_ROWS)
        self.translation_group.toggled.connect(self._on_group_toggled)
        self.services_layout.addWidget(self.translation_group)
        layout.addLayout(self.services_layout)
        self.setLayout(layout)
        self.set_language(self.lang)

    def _on_delete_clicked(self) -> None:
        self.remove_requested.emit()

    @staticmethod
    def _format_rate(value: Union[str, float], sep: str | None = None) -> str:
        return format_rate(value, sep)

    def create_service_group(self, service_name: str, rows: List[Dict]) -> QGroupBox:
        group = QGroupBox(tr(service_name, self.lang))
        group.setCheckable(True)
        group.setChecked(True)

        vbox = QVBoxLayout()

        rows = copy.deepcopy(rows)
        for r in rows:
            r['deleted'] = False

        table = QTableWidget(len(rows), 5)
        table.setHorizontalHeaderLabels([
            tr("Параметр", self.lang),
            tr("Ед-ца", self.lang),
            tr("Объем", self.lang),
            f"{tr('Ставка', self.lang)} ({self.currency_symbol})",
            f"{tr('Сумма', self.lang)} ({self.currency_symbol})",
        ])
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setSelectionMode(QAbstractItemView.ExtendedSelection)

        # ВАЖНО: никаких локальных скроллов — всё видно сразу
        table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        table.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        table.setWordWrap(False)

        base_rate_row = None
        for i, row_info in enumerate(rows):
            table.setItem(i, 0, QTableWidgetItem(tr(row_info["name"], self.lang)))
            unit_item = QTableWidgetItem(tr("слово", self.lang))
            table.setItem(i, 1, unit_item)
            table.setItem(i, 2, QTableWidgetItem("0"))

            rate_item = QTableWidgetItem("0")
            if not row_info["is_base"]:
                rate_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            else:
                if base_rate_row is None:
                    base_rate_row = i
            table.setItem(i, 3, rate_item)

            sum_item = QTableWidgetItem("0.00")
            sum_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            table.setItem(i, 4, sum_item)

        # Автоподгон ширин
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)

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
            add_act = menu.addAction(tr("Добавить строку", self.lang))
            del_act = menu.addAction(tr("Удалить", self.lang))
            restore_act = menu.addAction(tr("Восстановить", self.lang))

            fuzzy_menu = menu.addMenu(tr("Фаззи", self.lang))
            fuzzy_actions = {}
            for cfg in ServiceConfig.TRANSLATION_ROWS[1:]:
                act = fuzzy_menu.addAction(tr(cfg["name"], self.lang))
                fuzzy_actions[act] = cfg

            row_cfg = rows[row]

            selected_deleted_rows = sorted(
                {
                    index.row()
                    for index in table.selectedIndexes()
                    if 0 <= index.row() < len(rows)
                    and rows[index.row()].get("deleted")
                }
            )

            active_rows = sum(1 for r in rows if not r.get("deleted"))
            selected_rows = sorted(
                {
                    index.row()
                    for index in table.selectedIndexes()
                    if 0 <= index.row() < len(rows)
                    and not rows[index.row()].get("deleted")
                }
            )
            can_delete = False
            if selected_rows:
                can_delete = active_rows - len(selected_rows) >= 1
            else:
                can_delete = active_rows > 1 and not row_cfg.get("deleted")
            if not can_delete:
                del_act.setEnabled(False)

            can_restore = False
            if selected_deleted_rows:
                can_restore = True
            elif row_cfg.get("deleted"):
                can_restore = True
            if not can_restore:
                restore_act.setEnabled(False)
            action = menu.exec(table.mapToGlobal(pos))
            if action == add_act:
                self._add_row_after(table, rows, group, row)
            elif action == del_act:
                targets = selected_rows if selected_rows else [row]
                self._delete_rows(table, rows, group, targets)
            elif action == restore_act:
                targets = selected_deleted_rows if selected_deleted_rows else [row]
                self._restore_rows(table, rows, group, targets)
            elif action in fuzzy_actions:
                self._add_row_after(table, rows, group, row, fuzzy_actions[action])

        table.setContextMenuPolicy(Qt.CustomContextMenu)
        table.customContextMenuRequested.connect(show_menu)

        # Промежуточная сумма и меню модификаторов
        subtotal_layout = QHBoxLayout()
        subtotal_layout.setContentsMargins(0, 0, 0, 0)
        subtotal_label = QLabel()
        subtotal_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        subtotal_label.setObjectName("subtotal_label")
        subtotal_layout.addWidget(subtotal_label, 1)

        modifiers_button = QPushButton("⚙️")
        modifiers_button.setFlat(True)
        modifiers_button.setCursor(Qt.PointingHandCursor)
        modifiers_button.clicked.connect(self._show_modifiers_menu)
        subtotal_layout.addWidget(modifiers_button)

        vbox.addLayout(subtotal_layout)

        group.setLayout(vbox)

        # сохранить ссылки
        setattr(group, 'table', table)
        setattr(group, 'rows_config', rows)
        setattr(group, 'base_rate_row', base_rate_row)
        setattr(group, 'subtotal_label', subtotal_label)
        setattr(group, 'modifiers_button', modifiers_button)

        # начальный пересчёт + авто-раскрытие высоты
        self.update_rates_and_sums(table, rows, base_rate_row)
        self._fit_table_height(table)

        return group

    def _add_row_after(
        self,
        table: QTableWidget,
        rows: List[Dict],
        group: QGroupBox,
        row: int,
        row_cfg: Dict | None = None,
    ):
        base_rate_row = getattr(group, 'base_rate_row', None)
        insert_at = row + 1
        table.insertRow(insert_at)

        if row_cfg is None:
            name = tr("Новая строка", self.lang)
            new_cfg = {
                "name": "Новая строка",
                "is_base": False,
                "multiplier": 1.0,
                "deleted": False,
            }
        else:
            name = tr(row_cfg["name"], self.lang)
            new_cfg = {
                "name": row_cfg["name"],
                "key": row_cfg.get("key"),
                "is_base": row_cfg.get("is_base", False),
                "multiplier": row_cfg.get("multiplier", 1.0),
                "deleted": False,
            }

        rows.insert(insert_at, new_cfg)
        table.blockSignals(True)
        table.setItem(insert_at, 0, QTableWidgetItem(name))
        table.setItem(insert_at, 1, QTableWidgetItem(tr("слово", self.lang)))
        table.setItem(insert_at, 2, QTableWidgetItem("0"))
        rate_item = QTableWidgetItem("0")
        rate_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        table.setItem(insert_at, 3, rate_item)
        sum_item = QTableWidgetItem("0.00")
        sum_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        table.setItem(insert_at, 4, sum_item)
        table.blockSignals(False)
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
                if col == 3:
                    if rows[row]['is_base']:
                        item.setFlags(Qt.ItemIsEditable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                    else:
                        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                elif col == 4:
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                else:
                    item.setFlags(Qt.ItemIsEditable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            font = item.font()
            font.setStrikeOut(deleted)
            item.setFont(font)

    def _delete_row(self, table: QTableWidget, rows: List[Dict], group: QGroupBox, row: int):
        self._delete_rows(table, rows, group, [row])

    def _delete_rows(
        self,
        table: QTableWidget,
        rows: List[Dict],
        group: QGroupBox,
        targets: List[int],
    ):
        active = [idx for idx, cfg in enumerate(rows) if not cfg.get("deleted")]
        if len(active) <= 1:
            return
        removable = sorted(
            {
                r
                for r in targets
                if 0 <= r < len(rows) and not rows[r].get("deleted")
            }
        )
        if not removable:
            return
        max_remove = len(active) - 1
        if max_remove <= 0:
            return
        if len(removable) > max_remove:
            removable = removable[-max_remove:]
        base_rate_row = getattr(group, 'base_rate_row', None)
        for row in removable:
            self._set_row_deleted(table, rows, row, True)
            if base_rate_row == row:
                base_rate_row = None
        setattr(group, 'base_rate_row', base_rate_row)
        self.update_rates_and_sums(table, rows, base_rate_row)

    def _restore_row(self, table: QTableWidget, rows: List[Dict], group: QGroupBox, row: int):
        self._restore_rows(table, rows, group, [row])

    def _restore_rows(self, table: QTableWidget, rows: List[Dict], group: QGroupBox, targets: List[int]):
        base_rate_row = getattr(group, 'base_rate_row', None)
        restored_any = False
        for row in sorted({r for r in targets if 0 <= r < len(rows) and rows[r].get('deleted')}):
            self._set_row_deleted(table, rows, row, False)
            restored_any = True
            if rows[row]['is_base']:
                base_rate_row = row
        if restored_any:
            setattr(group, 'base_rate_row', base_rate_row)
            self.update_rates_and_sums(table, rows, base_rate_row)

    # ---------------- Logic ----------------
    def set_only_new_and_repeats_mode(self, enabled: bool):
        """Переключает режим отображения: 4 строки либо только новые/повторы."""
        group = self.translation_group
        table: QTableWidget = group.table
        rows = group.rows_config
        base_rate_row = getattr(group, 'base_rate_row', 0)

        if enabled and not self.only_new_repeats_mode:
            # сохраняем текущие значения для восстановления
            self._backup_volumes = [
                table.item(i, 2).text() if table.item(i, 2) else "0"
                for i in range(table.rowCount())
            ]
            self._backup_rates = [
                table.item(i, 3).text() if table.item(i, 3) else "0"
                for i in range(table.rowCount())
            ]

            total_new = 0.0
            for idx in range(min(3, table.rowCount())):
                try:
                    total_new += _to_float(self._backup_volumes[idx])
                except ValueError:
                    pass
            if table.item(0, 2):
                table.item(0, 2).setText(str(total_new))

            if table.rowCount() > 1:
                table.setRowHidden(1, True)
                rows[1]['deleted'] = True
            if table.rowCount() > 2:
                table.setRowHidden(2, True)
                rows[2]['deleted'] = True
            self.only_new_repeats_mode = True

        elif not enabled and self.only_new_repeats_mode:
            for idx in range(min(len(self._backup_volumes), table.rowCount())):
                if table.item(idx, 2):
                    table.item(idx, 2).setText(self._backup_volumes[idx])
                if table.item(idx, 3):
                    table.item(idx, 3).setText(self._backup_rates[idx])
            if table.rowCount() > 1:
                table.setRowHidden(1, False)
                rows[1]['deleted'] = False
            if table.rowCount() > 2:
                table.setRowHidden(2, False)
                rows[2]['deleted'] = False
            self.only_new_repeats_mode = False

        self.update_rates_and_sums(table, rows, base_rate_row)
        self._fit_table_height(table)

    def toggle_only_new_and_repeats(self):
        self.set_only_new_and_repeats_mode(not self.only_new_repeats_mode)

    def _fit_table_height(self, table: QTableWidget):
        """Делает таблицу фиксированной высоты по всем строкам (без внутреннего скролла)."""
        header_h = table.horizontalHeader().height()
        rows_h = sum(table.rowHeight(r) for r in range(table.rowCount()))
        frame = table.frameWidth() * 2
        # небольшой запас под сетку/паддинги
        total = header_h + rows_h + frame + 2
        table.setMinimumHeight(total)
        table.setMaximumHeight(total)

    def _on_group_toggled(self, checked: bool):
        group = getattr(self, "translation_group", None)
        if not group:
            return
        for attr in (
            "subtotal_label",
            "modifiers_button",
        ):
            widget = getattr(group, attr, None)
            if widget:
                widget.setEnabled(checked)
        self._update_discount_label()
        self.subtotal_changed.emit(self.get_subtotal())

    def _update_discount_label(self):
        group = getattr(self, "translation_group", None)
        if not group:
            return
        subtotal_label: QLabel | None = getattr(group, "subtotal_label", None)
        if not subtotal_label:
            return

        base_total = self._subtotal
        discount_percent = self._discount_percent if group.isChecked() else 0.0
        markup_percent = self._markup_percent if group.isChecked() else 0.0
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

        prefix_amount = base_total if group.isChecked() else final_total
        prefix = f"{tr('Промежуточная сумма', self.lang)}: {self._format_currency(prefix_amount)}"
        if parts:
            subtotal_label.setText(
                f"{prefix} {' '.join(parts)} = {self._format_currency(final_total)}"
            )
        else:
            subtotal_label.setText(prefix)

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

    def _show_modifiers_menu(self):
        group = getattr(self, "translation_group", None)
        if not group or not group.isEnabled():
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
            # При конфликте показываем скидку приоритетно
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

        button: QPushButton | None = getattr(group, "modifiers_button", None)
        if button:
            menu.exec(button.mapToGlobal(button.rect().bottomLeft()))
        else:
            menu.exec(self.mapToGlobal(self.rect().center()))

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

    def set_language(self, lang: str):
        self.lang = lang
        self.title_label.setText(f"{tr('Языковая пара', lang)}:")
        self.title_edit.setText(self.pair_name)
        if getattr(self, "delete_button", None):
            tooltip = tr("Удалить языковую пару", lang)
            self.delete_button.setToolTip(tooltip)
            self.delete_button.setAccessibleName(tooltip)
        group = self.translation_group
        group.setTitle(tr("Перевод", lang))
        table: QTableWidget = group.table
        table.setHorizontalHeaderLabels([
            tr("Параметр", lang),
            tr("Ед-ца", lang),
            tr("Объем", lang),
            f"{tr('Ставка', lang)} ({self.currency_symbol})",
            f"{tr('Сумма', lang)} ({self.currency_symbol})",
        ])
        rows = group.rows_config
        for i, row_info in enumerate(rows):
            item = table.item(i, 0)
            if item:
                item.setText(tr(row_info["name"], lang))
        self.update_rates_and_sums(table, rows, getattr(group, 'base_rate_row'))
        self._update_reports_label()

    def set_pair_name(self, name: str):
        self.pair_name = name
        self.title_edit.setText(name)

    def _on_title_edit(self):
        new_name = self.title_edit.text().strip()
        if new_name:
            self.pair_name = new_name
            self.name_changed.emit(new_name)

    def update_rates_and_sums(self, table: QTableWidget, rows: List[Dict], base_rate_row: int):
        try:
            base_rate = 0.0
            base_sep = "."
            if base_rate_row is not None and rows[base_rate_row].get('deleted'):
                base_rate_row = None
            if base_rate_row is not None and table.item(base_rate_row, 3):
                base_text = table.item(base_rate_row, 3).text()
                if self.lang == "en":
                    base_sep = "."
                else:
                    base_sep = "," if "," in base_text else "."
                base_rate = _to_float(base_text)
                table.blockSignals(True)
                table.item(base_rate_row, 3).setText(self._format_rate(base_text, base_sep))
                table.blockSignals(False)

            subtotal = 0.0
            for row in range(min(table.rowCount(), len(rows))):
                row_cfg = rows[row]
                if row_cfg.get('deleted'):
                    if table.item(row, 4):
                        table.blockSignals(True)
                        table.item(row, 4).setText("0.00")
                        table.blockSignals(False)
                    continue

                # авто-ставки для непервой строки
                if not row_cfg["is_base"] and base_rate_row is not None:
                    auto_rate = base_rate * row_cfg["multiplier"]
                    if table.item(row, 3):
                        table.blockSignals(True)
                        table.item(row, 3).setText(self._format_rate(auto_rate, base_sep))
                        table.blockSignals(False)

                volume = _to_float(table.item(row, 2).text() if table.item(row, 2) else "0")
                rate_item = table.item(row, 3)
                rate_text = rate_item.text() if rate_item else "0"
                if self.lang == "en":
                    sep = "."
                else:
                    sep = "," if "," in rate_text else "."
                rate = _to_float(rate_text)
                table.blockSignals(True)
                rate_item.setText(self._format_rate(rate_text, sep))
                table.blockSignals(False)
                total = volume * rate
                if table.item(row, 4):
                    table.blockSignals(True)
                    table.item(row, 4).setText(format_amount(total, self.lang))
                    table.blockSignals(False)
                subtotal += total

            # обновить «Промежуточную сумму»
            self._subtotal = subtotal
            self._update_discount_label()
            self.subtotal_changed.emit(self.get_subtotal())

            # после любых изменений гарантируем отсутствие локального скролла
            self._fit_table_height(table)

        except (ValueError, AttributeError):
            pass

    def refresh_layout(self):
        """Публичный хук для внешнего кода: раскрыть таблицы по высоте ещё раз."""
        if hasattr(self.translation_group, "table"):
            self._fit_table_height(self.translation_group.table)

    def get_subtotal(self) -> float:
        if not self.translation_group.isChecked():
            return 0.0
        base = self._subtotal
        discount_amount = base * (self._discount_percent / 100.0)
        markup_amount = base * (self._markup_percent / 100.0)
        return base - discount_amount + markup_amount

    def get_discount_amount(self) -> float:
        if not self.translation_group.isChecked():
            return 0.0
        return self._subtotal * (self._discount_percent / 100.0)

    def get_markup_amount(self) -> float:
        if not self.translation_group.isChecked():
            return 0.0
        return self._subtotal * (self._markup_percent / 100.0)

    # ---------------- Data ----------------
    def get_data(self) -> Dict[str, Any]:
        data = {
            "pair_name": self.pair_name,
            "services": {},
            "report_sources": self.report_sources(),
            "only_new_repeats": self.only_new_repeats_mode,
            "discount_percent": self.get_discount_percent(),
            "discount_amount": self.get_discount_amount(),
            "markup_percent": self.get_markup_percent(),
            "markup_amount": self.get_markup_amount(),
        }
        if self.translation_group.isChecked():
            data["services"]["translation"] = self._get_table_data(self.translation_group.table)
        return data

    def _get_table_data(self, table: QTableWidget) -> List[Dict[str, Any]]:
        out = []
        rows_cfg = self.translation_group.rows_config
        for row in range(min(table.rowCount(), len(rows_cfg))):
            row_cfg = rows_cfg[row]
            out.append({
                "key": row_cfg.get("key"),
                "name": row_cfg.get("name"),
                "parameter": table.item(row, 0).text() if table.item(row, 0) else "",
                "unit": table.item(row, 1).text() if table.item(row, 1) else "",
                "volume": _to_float(table.item(row, 2).text() if table.item(row, 2) else "0"),
                "rate": _to_float(table.item(row, 3).text() if table.item(row, 3) else "0"),
                "total": _to_float(table.item(row, 4).text() if table.item(row, 4) else "0"),
                "is_base": row_cfg.get("is_base", False),
                "multiplier": row_cfg.get("multiplier"),
                "deleted": row_cfg.get("deleted", False),
            })
        return out

    def load_table_data(self, data: List[Dict[str, Any]]):
        group = self.translation_group
        table = group.table
        rows = group.rows_config
        base_rate_row = None

        for idx in range(table.rowCount()):
            rows[idx]['deleted'] = False
            table.setRowHidden(idx, False)

        data_by_key = {
            row.get("key"): row for row in data if row.get("key") is not None
        }
        assigned_rows: List[Dict[str, Any] | None] = [None] * table.rowCount()
        used_keys = set()

        for idx, row_cfg in enumerate(rows):
            key = row_cfg.get("key")
            if key in data_by_key:
                assigned_rows[idx] = data_by_key[key]
                used_keys.add(key)

        remaining = [row for row in data if row.get("key") not in used_keys]

        for idx in range(len(assigned_rows)):
            if assigned_rows[idx] is None and remaining:
                assigned_rows[idx] = remaining.pop(0)

        while remaining:
            r = table.rowCount()
            table.insertRow(r)
            table.setItem(r, 0, QTableWidgetItem(tr("Новая строка", self.lang)))
            table.setItem(r, 1, QTableWidgetItem(tr("слово", self.lang)))
            table.setItem(r, 2, QTableWidgetItem("0"))
            rate_item = QTableWidgetItem("0")
            rate_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            table.setItem(r, 3, rate_item)
            sum_item = QTableWidgetItem("0.00")
            sum_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            table.setItem(r, 4, sum_item)
            rows.append({"name": "Новая строка", "is_base": False, "multiplier": 1.0, "deleted": False})
            assigned_rows.append(remaining.pop(0))

        for row, row_data in enumerate(assigned_rows):
            if row_data is None:
                continue
            table.item(row, 0).setText(row_data.get("parameter", ""))
            unit_value = row_data.get("unit") or tr("слово", self.lang)
            table.item(row, 1).setText(str(unit_value))
            table.item(row, 2).setText(str(row_data.get("volume", 0)))
            sep = "." if self.lang == "en" else None
            table.item(row, 3).setText(self._format_rate(row_data.get('rate', 0), sep))
            table.item(row, 4).setText(format_amount(row_data.get('total', 0), self.lang))
            rows[row]["is_base"] = row_data.get("is_base", rows[row].get("is_base", False))
            rows[row]["multiplier"] = row_data.get("multiplier", rows[row].get("multiplier", 1.0))
            rows[row]["key"] = row_data.get("key", rows[row].get("key"))
            rows[row]["name"] = row_data.get("name", rows[row].get("name"))
            deleted_flag = row_data.get("deleted", False)
            rows[row]["deleted"] = deleted_flag
            self._set_row_deleted(table, rows, row, deleted_flag)
            if rows[row].get("is_base"):
                base_rate_row = row

        setattr(group, 'base_rate_row', base_rate_row)
        self.update_rates_and_sums(table, rows, base_rate_row)
        self._fit_table_height(table)

    def set_report_sources(self, sources: Iterable[str], replace: bool = False) -> None:
        if replace:
            self._report_sources = []
        existing = set(self._report_sources)
        for name in sources:
            cleaned = name.strip()
            if cleaned and cleaned not in existing:
                self._report_sources.append(cleaned)
                existing.add(cleaned)
        self._update_reports_label()

    def clear_report_sources(self) -> None:
        self._report_sources = []
        self._update_reports_label()

    def report_sources(self) -> List[str]:
        return list(self._report_sources)

    def _update_reports_label(self) -> None:
        if not self._report_sources:
            self.reports_label.hide()
            self.reports_label.setToolTip("")
            return
        prefix = tr("Отчёты", self.lang)
        joined = ", ".join(self._report_sources)
        self.reports_label.setText(f"{prefix}: {joined}")
        self.reports_label.setToolTip(joined)
        self.reports_label.show()

    def set_basic_rate(self, value: float) -> None:
        """Set base translation rate and update dependent rows."""
        group = self.translation_group
        table = getattr(group, 'table', None)
        rows = getattr(group, 'rows_config', None)
        base_row = getattr(group, 'base_rate_row', None)
        if table is None or rows is None or base_row is None:
            return
        sep = "." if self.lang == "en" else None
        item = table.item(base_row, 3)
        if item is None:
            item = QTableWidgetItem()
            table.setItem(base_row, 3, item)
        item.setText(self._format_rate(value, sep))
        self.update_rates_and_sums(table, rows, base_row)

    def set_currency(self, symbol: str, code: str):
        self.currency_symbol = symbol
        self.currency_code = code
        if hasattr(self.translation_group, 'table'):
            self.translation_group.table.setHorizontalHeaderLabels([
                tr("Параметр", self.lang),
                tr("Ед-ца", self.lang),
                tr("Объем", self.lang),
                f"{tr('Ставка', self.lang)} ({symbol})",
                f"{tr('Сумма', self.lang)} ({symbol})",
            ])
            self.update_rates_and_sums(
                self.translation_group.table,
                self.translation_group.rows_config,
                getattr(self.translation_group, 'base_rate_row')
            )

    def convert_rates(self, multiplier: float):
        """Multiply all rate values by *multiplier* and update totals."""
        group = getattr(self, 'translation_group', None)
        if not group or not hasattr(group, 'table'):
            return
        table: QTableWidget = group.table
        rows = group.rows_config
        for row in range(table.rowCount()):
            if rows[row].get('deleted'):
                continue
            item = table.item(row, 3)
            if item is None:
                continue
            rate = _to_float(item.text())
            sep = '.' if self.lang == 'en' else ','
            item.setText(self._format_rate(rate * multiplier, sep))
        self.update_rates_and_sums(table, rows, getattr(group, 'base_rate_row'))
