from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from logic.legal_entities import (
    add_or_update_legal_entity,
    ensure_user_template_copy,
    export_legal_entity_template,
    list_legal_entities_detailed,
    remove_user_legal_entity,
)
from logic.service_config import ServiceConfig
from logic.translation_config import tr
from logic.user_config import load_languages, reset_languages, save_languages


class SettingsDialog(QDialog):
    settings_updated = Signal()

    def __init__(self, parent: Optional[QWidget] = None, lang: str = "ru") -> None:
        super().__init__(parent)
        self.lang = lang
        self.setWindowTitle(tr("Настройки", lang))
        self.setModal(True)

        layout = QVBoxLayout(self)

        self.tabs = QTabWidget(self)
        self.legal_tab = LegalEntitiesTab(lang, self)
        self.languages_tab = LanguagesTab(lang, self)
        self.tariffs_tab = TariffConfigTab(lang, self)

        for tab in (self.legal_tab, self.languages_tab, self.tariffs_tab):
            tab.changed.connect(self._emit_updated)

        self.tabs.addTab(self.legal_tab, tr("Юрлица", lang))
        self.tabs.addTab(self.languages_tab, tr("Языки", lang))
        self.tabs.addTab(self.tariffs_tab, tr("Тарифы", lang))
        layout.addWidget(self.tabs)

        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(self.reject)
        buttons.accepted.connect(self.accept)
        buttons.button(QDialogButtonBox.Close).setText(tr("Закрыть", lang))
        layout.addWidget(buttons)

    def _emit_updated(self) -> None:
        self.settings_updated.emit()


class LegalEntitiesTab(QWidget):
    changed = Signal()

    def __init__(self, lang: str, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.lang = lang
        self._setup_ui()
        self.refresh()

    def _setup_ui(self) -> None:
        layout = QVBoxLayout(self)

        info_label = QLabel(
            tr(
                "Загрузите новые шаблоны или экспортируйте существующие Excel-файлы.",
                self.lang,
            )
        )
        info_label.setWordWrap(True)
        layout.addWidget(info_label)

        self.table = QTableWidget(0, 3, self)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setHorizontalHeaderLabels(
            [
                tr("Название", self.lang),
                tr("Файл", self.lang),
                tr("Источник", self.lang),
            ]
        )
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        layout.addWidget(self.table)
        self.table.itemSelectionChanged.connect(self._update_buttons_state)

        buttons_layout = QHBoxLayout()
        self.add_btn = QPushButton(tr("Добавить", self.lang))
        self.add_btn.clicked.connect(self.add_template)
        buttons_layout.addWidget(self.add_btn)

        self.replace_btn = QPushButton(tr("Заменить", self.lang))
        self.replace_btn.clicked.connect(self.replace_template)
        buttons_layout.addWidget(self.replace_btn)

        self.export_btn = QPushButton(tr("Экспортировать", self.lang))
        self.export_btn.clicked.connect(self.export_template)
        buttons_layout.addWidget(self.export_btn)

        self.remove_btn = QPushButton(tr("Удалить", self.lang))
        self.remove_btn.clicked.connect(self.remove_template)
        buttons_layout.addWidget(self.remove_btn)

        self.refresh_btn = QPushButton(tr("Обновить", self.lang))
        self.refresh_btn.clicked.connect(self.refresh)
        buttons_layout.addWidget(self.refresh_btn)

        buttons_layout.addStretch()
        layout.addLayout(buttons_layout)

    def refresh(self) -> None:
        details = list_legal_entities_detailed()
        self.table.setRowCount(len(details))
        for row, entry in enumerate(details):
            name_item = QTableWidgetItem(entry.get("name", ""))
            name_item.setData(Qt.UserRole, entry)
            file_path = entry.get("resolved_template") or entry.get("template") or ""
            file_item = QTableWidgetItem(file_path)
            source_text = (
                tr("Пользовательский", self.lang)
                if entry.get("source") == "user"
                else tr("Стандартный", self.lang)
            )
            source_item = QTableWidgetItem(source_text)
            self.table.setItem(row, 0, name_item)
            self.table.setItem(row, 1, file_item)
            self.table.setItem(row, 2, source_item)
        self._update_buttons_state()

    def _update_buttons_state(self) -> None:
        entry = self._selected_entry()
        has_entry = entry is not None
        self.replace_btn.setEnabled(has_entry)
        self.export_btn.setEnabled(has_entry)
        self.remove_btn.setEnabled(has_entry and entry.get("source") == "user")

    def _selected_entry(self) -> Optional[Dict[str, Any]]:
        row = self.table.currentRow()
        if row < 0:
            return None
        item = self.table.item(row, 0)
        if not item:
            return None
        data = item.data(Qt.UserRole)
        if isinstance(data, dict):
            return data
        return None

    def _ask_for_name(self, default: str = "") -> Optional[str]:
        text, ok = QInputDialog.getText(
            self,
            tr("Название шаблона", self.lang),
            tr("Введите отображаемое название шаблона", self.lang),
            text=default,
        )
        if not ok:
            return None
        name = text.strip()
        return name or None

    def add_template(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            tr("Выберите шаблон Excel", self.lang),
            "",
            "Excel (*.xlsx *.xls);;" + tr("Все файлы", self.lang) + " (*)",
        )
        if not file_path:
            return
        name = self._ask_for_name(Path(file_path).stem)
        if not name:
            return
        existing = self._find_entry(name)
        if existing:
            answer = QMessageBox.question(
                self,
                tr("Подтверждение", self.lang),
                tr("Перезаписать существующий шаблон?", self.lang),
            )
            if answer != QMessageBox.Yes:
                return
        destination = ensure_user_template_copy(Path(file_path), name)
        add_or_update_legal_entity(name, str(destination))
        self.changed.emit()
        self.refresh()

    def replace_template(self) -> None:
        entry = self._selected_entry()
        if not entry:
            return
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            tr("Выберите шаблон Excel", self.lang),
            "",
            "Excel (*.xlsx *.xls);;" + tr("Все файлы", self.lang) + " (*)",
        )
        if not file_path:
            return
        name = entry.get("name") or Path(file_path).stem
        destination = ensure_user_template_copy(Path(file_path), name)
        add_or_update_legal_entity(name, str(destination))
        self.changed.emit()
        self.refresh()

    def export_template(self) -> None:
        entry = self._selected_entry()
        if not entry:
            return
        template_path = entry.get("resolved_template") or entry.get("template")
        default_name = Path(template_path or "template").name or "template.xlsx"
        dest_path, _ = QFileDialog.getSaveFileName(
            self,
            tr("Сохранить шаблон", self.lang),
            default_name,
            "Excel (*.xlsx *.xls);;" + tr("Все файлы", self.lang) + " (*)",
        )
        if not dest_path:
            return
        destination = Path(dest_path)
        if destination.suffix.lower() not in {".xlsx", ".xls"}:
            destination = destination.with_suffix(".xlsx")
        success = export_legal_entity_template(entry.get("name", ""), destination)
        if success:
            QMessageBox.information(
                self,
                tr("Экспорт завершён", self.lang),
                tr("Шаблон успешно сохранён.", self.lang),
            )
        else:
            QMessageBox.warning(
                self,
                tr("Ошибка", self.lang),
                tr("Не удалось сохранить шаблон.", self.lang),
            )

    def remove_template(self) -> None:
        entry = self._selected_entry()
        if not entry or entry.get("source") != "user":
            return
        answer = QMessageBox.question(
            self,
            tr("Подтверждение", self.lang),
            tr("Удалить пользовательский шаблон?", self.lang),
        )
        if answer != QMessageBox.Yes:
            return
        remove_user_legal_entity(entry.get("name", ""))
        self.changed.emit()
        self.refresh()

    def _find_entry(self, name: str) -> Optional[Dict[str, Any]]:
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            if item and item.text().strip().lower() == name.strip().lower():
                data = item.data(Qt.UserRole)
                if isinstance(data, dict):
                    return data
        return None


class LanguagesTab(QWidget):
    changed = Signal()

    def __init__(self, lang: str, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.lang = lang
        self._setup_ui()
        self.refresh()

    def _setup_ui(self) -> None:
        layout = QVBoxLayout(self)
        self.table = QTableWidget(0, 2, self)
        self.table.setHorizontalHeaderLabels(
            [tr("Название RU", self.lang), tr("Название EN", self.lang)]
        )
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        layout.addWidget(self.table)

        buttons_layout = QHBoxLayout()
        self.add_btn = QPushButton(tr("Добавить", self.lang))
        self.add_btn.clicked.connect(self.add_row)
        buttons_layout.addWidget(self.add_btn)

        self.remove_btn = QPushButton(tr("Удалить", self.lang))
        self.remove_btn.clicked.connect(self.remove_selected)
        buttons_layout.addWidget(self.remove_btn)

        self.save_btn = QPushButton(tr("Сохранить", self.lang))
        self.save_btn.clicked.connect(self.save_changes)
        buttons_layout.addWidget(self.save_btn)

        self.reset_btn = QPushButton(tr("Сбросить", self.lang))
        self.reset_btn.clicked.connect(self.reset_default)
        buttons_layout.addWidget(self.reset_btn)

        buttons_layout.addStretch()
        layout.addLayout(buttons_layout)

    def refresh(self) -> None:
        data = load_languages()
        self.table.setRowCount(len(data))
        for row, entry in enumerate(data):
            ru_item = QTableWidgetItem(entry.get("ru", ""))
            en_item = QTableWidgetItem(entry.get("en", ""))
            for item in (ru_item, en_item):
                item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEditable | Qt.ItemIsEnabled)
            self.table.setItem(row, 0, ru_item)
            self.table.setItem(row, 1, en_item)

    def add_row(self) -> None:
        row = self.table.rowCount()
        self.table.insertRow(row)
        for col in range(2):
            item = QTableWidgetItem("")
            item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEditable | Qt.ItemIsEnabled)
            self.table.setItem(row, col, item)

    def remove_selected(self) -> None:
        rows = sorted({index.row() for index in self.table.selectedIndexes()}, reverse=True)
        for row in rows:
            self.table.removeRow(row)

    def save_changes(self) -> None:
        result: List[Dict[str, str]] = []
        for row in range(self.table.rowCount()):
            ru_item = self.table.item(row, 0)
            en_item = self.table.item(row, 1)
            ru = ru_item.text().strip() if ru_item else ""
            en = en_item.text().strip() if en_item else ""
            if not (ru or en):
                continue
            result.append({"ru": ru, "en": en})
        if not result:
            QMessageBox.warning(
                self,
                tr("Ошибка", self.lang),
                tr("Список языков не может быть пустым.", self.lang),
            )
            return
        if save_languages(result):
            QMessageBox.information(
                self,
                tr("Сохранено", self.lang),
                tr("Справочник языков обновлён.", self.lang),
            )
            self.changed.emit()
        else:
            QMessageBox.warning(
                self,
                tr("Ошибка", self.lang),
                tr("Не удалось сохранить языки.", self.lang),
            )

    def reset_default(self) -> None:
        answer = QMessageBox.question(
            self,
            tr("Подтверждение", self.lang),
            tr("Сбросить справочник к настройкам по умолчанию?", self.lang),
        )
        if answer != QMessageBox.Yes:
            return
        if reset_languages():
            self.refresh()
            QMessageBox.information(
                self,
                tr("Сброшено", self.lang),
                tr("Справочник восстановлен по умолчанию.", self.lang),
            )
            self.changed.emit()
        else:
            QMessageBox.warning(
                self,
                tr("Ошибка", self.lang),
                tr("Не удалось сбросить языки.", self.lang),
            )


class TariffConfigTab(QWidget):
    changed = Signal()

    def __init__(self, lang: str, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.lang = lang
        self._setup_ui()
        self._load_from_config()

    def _setup_ui(self) -> None:
        layout = QVBoxLayout(self)

        translation_group = QGroupBox(tr("Перевод", self.lang), self)
        translation_layout = QVBoxLayout(translation_group)

        self.translation_table = QTableWidget(0, 4, translation_group)
        self.translation_table.setHorizontalHeaderLabels(
            [
                tr("Ключ", self.lang),
                tr("Название", self.lang),
                tr("Процент", self.lang),
                tr("Базовая", self.lang),
            ]
        )
        self.translation_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        header = self.translation_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        translation_layout.addWidget(self.translation_table)

        translation_buttons = QHBoxLayout()
        self.add_row_btn = QPushButton(tr("Добавить", self.lang))
        self.add_row_btn.clicked.connect(self.add_translation_row)
        translation_buttons.addWidget(self.add_row_btn)

        self.remove_row_btn = QPushButton(tr("Удалить", self.lang))
        self.remove_row_btn.clicked.connect(self.remove_translation_row)
        translation_buttons.addWidget(self.remove_row_btn)

        translation_buttons.addStretch()
        translation_layout.addLayout(translation_buttons)

        layout.addWidget(translation_group)

        services_group = QGroupBox(tr("Дополнительные услуги", self.lang), self)
        services_layout = QVBoxLayout(services_group)

        self.services_tree = QTreeWidget(services_group)
        self.services_tree.setColumnCount(3)
        self.services_tree.setHeaderLabels(
            [
                tr("Категория / Услуга", self.lang),
                tr("Процент", self.lang),
                tr("Базовая", self.lang),
            ]
        )
        self.services_tree.header().setSectionResizeMode(0, QHeaderView.Stretch)
        self.services_tree.header().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.services_tree.header().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        services_layout.addWidget(self.services_tree)

        services_buttons = QHBoxLayout()
        self.add_category_btn = QPushButton(tr("Добавить категорию", self.lang))
        self.add_category_btn.clicked.connect(self.add_category)
        services_buttons.addWidget(self.add_category_btn)

        self.add_service_btn = QPushButton(tr("Добавить услугу", self.lang))
        self.add_service_btn.clicked.connect(self.add_service)
        services_buttons.addWidget(self.add_service_btn)

        self.remove_service_btn = QPushButton(tr("Удалить", self.lang))
        self.remove_service_btn.clicked.connect(self.remove_selected)
        services_buttons.addWidget(self.remove_service_btn)

        services_buttons.addStretch()
        services_layout.addLayout(services_buttons)

        layout.addWidget(services_group)

        actions = QHBoxLayout()
        self.save_btn = QPushButton(tr("Сохранить", self.lang))
        self.save_btn.clicked.connect(self.save_config)
        actions.addWidget(self.save_btn)

        self.reset_btn = QPushButton(tr("Сбросить", self.lang))
        self.reset_btn.clicked.connect(self.reset_defaults)
        actions.addWidget(self.reset_btn)

        actions.addStretch()
        layout.addLayout(actions)

    def _load_from_config(self) -> None:
        config = ServiceConfig.get_config()
        self._load_translation_rows(config.get("translation_rows", []))
        self._load_additional_services(config.get("additional_services", {}))

    def _load_translation_rows(self, rows: List[Dict[str, Any]]) -> None:
        self.translation_table.setRowCount(0)
        for row in rows:
            idx = self.translation_table.rowCount()
            self.translation_table.insertRow(idx)

            key_item = QTableWidgetItem(row.get("key") or "")
            key_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEditable | Qt.ItemIsEnabled)
            self.translation_table.setItem(idx, 0, key_item)

            name_item = QTableWidgetItem(row.get("name") or "")
            name_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEditable | Qt.ItemIsEnabled)
            self.translation_table.setItem(idx, 1, name_item)

            spin = self._create_percent_spin(row.get("multiplier", 1.0) * 100)
            self.translation_table.setCellWidget(idx, 2, spin)

            base_item = QTableWidgetItem()
            base_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            base_item.setCheckState(Qt.Checked if row.get("is_base") else Qt.Unchecked)
            self.translation_table.setItem(idx, 3, base_item)

        if self.translation_table.rowCount() == 0:
            self.add_translation_row()

    def _create_percent_spin(self, value: float = 100.0) -> QDoubleSpinBox:
        spin = QDoubleSpinBox(self)
        spin.setRange(0.0, 1000.0)
        spin.setDecimals(2)
        spin.setSuffix(" %")
        spin.setValue(max(0.0, value))
        return spin

    def add_translation_row(self) -> None:
        idx = self.translation_table.rowCount()
        self.translation_table.insertRow(idx)
        for col in range(2):
            item = QTableWidgetItem("")
            item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEditable | Qt.ItemIsEnabled)
            self.translation_table.setItem(idx, col, item)
        self.translation_table.setCellWidget(idx, 2, self._create_percent_spin())
        base_item = QTableWidgetItem()
        base_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        base_item.setCheckState(Qt.Unchecked)
        self.translation_table.setItem(idx, 3, base_item)

    def remove_translation_row(self) -> None:
        rows = sorted({index.row() for index in self.translation_table.selectedIndexes()}, reverse=True)
        if not rows:
            return
        if len(rows) >= self.translation_table.rowCount():
            QMessageBox.warning(
                self,
                tr("Ошибка", self.lang),
                tr("Должна остаться хотя бы одна строка.", self.lang),
            )
            return
        for row in rows:
            self.translation_table.removeRow(row)

    def _load_additional_services(self, services: Dict[str, List[Dict[str, Any]]]) -> None:
        self.services_tree.clear()
        for category, rows in services.items():
            cat_item = QTreeWidgetItem([category])
            cat_item.setFlags(cat_item.flags() | Qt.ItemIsEditable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.services_tree.addTopLevelItem(cat_item)
            for row in rows:
                child = QTreeWidgetItem([row.get("name", "")])
                child.setFlags(child.flags() | Qt.ItemIsEditable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                spin = self._create_percent_spin(row.get("multiplier", 1.0) * 100)
                self.services_tree.setItemWidget(child, 1, spin)
                child.setCheckState(2, Qt.Checked if row.get("is_base") else Qt.Unchecked)
                cat_item.addChild(child)
            cat_item.setExpanded(True)

    def add_category(self) -> None:
        name, ok = QInputDialog.getText(
            self,
            tr("Новая категория", self.lang),
            tr("Введите название категории", self.lang),
        )
        if not ok:
            return
        cleaned = name.strip()
        if not cleaned:
            return
        cat_item = QTreeWidgetItem([cleaned])
        cat_item.setFlags(cat_item.flags() | Qt.ItemIsEditable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        self.services_tree.addTopLevelItem(cat_item)
        cat_item.setExpanded(True)

    def add_service(self) -> None:
        item = self.services_tree.currentItem()
        if item is None:
            QMessageBox.information(
                self,
                tr("Информация", self.lang),
                tr("Сначала выберите категорию.", self.lang),
            )
            return
        if item.parent() is not None:
            parent = item.parent()
        else:
            parent = item
        service_item = QTreeWidgetItem([tr("Новая услуга", self.lang)])
        service_item.setFlags(service_item.flags() | Qt.ItemIsEditable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        spin = self._create_percent_spin(100.0)
        parent.addChild(service_item)
        self.services_tree.setItemWidget(service_item, 1, spin)
        service_item.setCheckState(2, Qt.Unchecked)
        parent.setExpanded(True)

    def remove_selected(self) -> None:
        item = self.services_tree.currentItem()
        if item is None:
            return
        parent = item.parent()
        if parent is None:
            index = self.services_tree.indexOfTopLevelItem(item)
            if index >= 0:
                self.services_tree.takeTopLevelItem(index)
        else:
            parent.removeChild(item)

    def _collect_translation_rows(self) -> Optional[List[Dict[str, Any]]]:
        rows: List[Dict[str, Any]] = []
        for row in range(self.translation_table.rowCount()):
            key = self.translation_table.item(row, 0)
            name = self.translation_table.item(row, 1)
            spin = self.translation_table.cellWidget(row, 2)
            base_item = self.translation_table.item(row, 3)
            label = name.text().strip() if name else ""
            if not label:
                QMessageBox.warning(
                    self,
                    tr("Ошибка", self.lang),
                    tr("Название строки не может быть пустым.", self.lang),
                )
                return None
            multiplier = spin.value() / 100.0 if isinstance(spin, QDoubleSpinBox) else 0.0
            rows.append(
                {
                    "key": (key.text().strip() if key else "") or None,
                    "name": label,
                    "multiplier": multiplier,
                    "is_base": bool(base_item.checkState() == Qt.Checked) if base_item else False,
                }
            )
        if not any(row.get("is_base") for row in rows):
            rows[0]["is_base"] = True
        return rows

    def _collect_additional_services(self) -> Dict[str, List[Dict[str, Any]]]:
        result: Dict[str, List[Dict[str, Any]]] = {}
        for idx in range(self.services_tree.topLevelItemCount()):
            cat = self.services_tree.topLevelItem(idx)
            name = cat.text(0).strip()
            if not name:
                continue
            services: List[Dict[str, Any]] = []
            for j in range(cat.childCount()):
                child = cat.child(j)
                label = child.text(0).strip()
                if not label:
                    continue
                spin = self.services_tree.itemWidget(child, 1)
                multiplier = spin.value() / 100.0 if isinstance(spin, QDoubleSpinBox) else 0.0
                services.append(
                    {
                        "name": label,
                        "multiplier": multiplier,
                        "is_base": child.checkState(2) == Qt.Checked,
                    }
                )
            if services:
                result[name] = services
        return result

    def save_config(self) -> None:
        rows = self._collect_translation_rows()
        if rows is None:
            return
        services = self._collect_additional_services()
        try:
            ServiceConfig.save_config(rows, services)
        except Exception:
            QMessageBox.warning(
                self,
                tr("Ошибка", self.lang),
                tr("Не удалось сохранить конфигурацию.", self.lang),
            )
            return
        QMessageBox.information(
            self,
            tr("Сохранено", self.lang),
            tr("Конфигурация тарифов обновлена.", self.lang),
        )
        self.changed.emit()

    def reset_defaults(self) -> None:
        answer = QMessageBox.question(
            self,
            tr("Подтверждение", self.lang),
            tr("Вернуть конфигурацию по умолчанию?", self.lang),
        )
        if answer != QMessageBox.Yes:
            return
        ServiceConfig.reset_to_defaults()
        self._load_from_config()
        QMessageBox.information(
            self,
            tr("Сброшено", self.lang),
            tr("Настройки тарифов восстановлены.", self.lang),
        )
        self.changed.emit()
