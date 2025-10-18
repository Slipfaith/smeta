"""Settings dialog that allows managing templates and service coefficients."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QInputDialog,
    QPushButton,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from logic.legal_entities import (
    LegalEntityRecord,
    add_or_update_legal_entity,
    export_legal_entities_to_excel,
    list_legal_entities_with_sources,
    remove_user_legal_entity,
)
from logic.service_config import ServiceConfig
from logic.settings_store import load_settings, save_settings
from logic.translation_config import tr


@dataclass
class _StagedEntity:
    name: str
    source: str
    original_template: str
    metadata: Dict[str, Optional[float | bool | str]] = field(default_factory=dict)
    pending_template: Optional[str] = None
    is_new: bool = False
    to_delete: bool = False

    def effective_template(self) -> str:
        return self.pending_template or self.original_template


class SettingsDialog(QDialog):
    """Modal dialog that allows editing configurable settings."""

    def __init__(self, parent: QWidget | None = None, lang: str = "ru") -> None:
        super().__init__(parent)
        self.setModal(True)
        self.lang = lang
        self._settings = load_settings()
        self._initial_records = list_legal_entities_with_sources()
        self._staged_entities: Dict[str, _StagedEntity] = {
            name: _StagedEntity(
                name=name,
                source=record.source,
                original_template=record.template,
                metadata=dict(record.metadata or {}),
            )
            for name, record in self._initial_records.items()
        }

        self.setWindowTitle(tr("Настройки", self.lang))
        self.resize(960, 640)

        self._setup_ui()
        self._populate_translation_rows()
        self._populate_additional_services()
        self._populate_fuzzy_thresholds()
        self._populate_legal_entities()

    # ------------------------------------------------------------------ UI
    def _setup_ui(self) -> None:
        layout = QVBoxLayout(self)
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)

        self.translation_tab = QWidget()
        self._setup_translation_tab()
        self.tabs.addTab(self.translation_tab, tr("Тарифы", self.lang))

        self.additional_tab = QWidget()
        self._setup_additional_tab()
        self.tabs.addTab(self.additional_tab, tr("Доп. услуги", self.lang))

        self.legal_tab = QWidget()
        self._setup_legal_tab()
        self.tabs.addTab(self.legal_tab, tr("Юрлица", self.lang))

        self.buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        self.buttons.accepted.connect(self._on_accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)

    # Translation tab ---------------------------------------------------
    def _setup_translation_tab(self) -> None:
        layout = QVBoxLayout()

        thresholds_group = QGroupBox(tr("Fuzzy matching", self.lang))
        thresholds_layout = QGridLayout()
        thresholds_group.setLayout(thresholds_layout)

        self.new_words_spin = QSpinBox()
        self.new_words_spin.setRange(0, 100)
        self.other_spin = QSpinBox()
        self.other_spin.setRange(0, 100)

        thresholds_layout.addWidget(QLabel(tr("Новые слова", self.lang)), 0, 0)
        thresholds_layout.addWidget(self.new_words_spin, 0, 1)
        thresholds_layout.addWidget(QLabel(tr("Другие категории", self.lang)), 1, 0)
        thresholds_layout.addWidget(self.other_spin, 1, 1)

        layout.addWidget(thresholds_group)

        self.translation_table = QTableWidget(0, 4)
        self.translation_table.setHorizontalHeaderLabels(
            [
                tr("Ключ", self.lang),
                tr("Название", self.lang),
                tr("Коэффициент", self.lang),
                tr("Базовая строка", self.lang),
            ]
        )
        self.translation_table.horizontalHeader().setStretchLastSection(True)
        self.translation_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.translation_table.setSelectionMode(QAbstractItemView.ExtendedSelection)

        layout.addWidget(self.translation_table)

        buttons_layout = QHBoxLayout()
        self.add_translation_row_btn = QPushButton(tr("Добавить", self.lang))
        self.remove_translation_row_btn = QPushButton(tr("Удалить", self.lang))
        self.translation_status = QLabel()
        self.translation_status.setWordWrap(True)

        self.add_translation_row_btn.clicked.connect(self._add_translation_row)
        self.remove_translation_row_btn.clicked.connect(self._remove_translation_rows)

        buttons_layout.addWidget(self.add_translation_row_btn)
        buttons_layout.addWidget(self.remove_translation_row_btn)
        buttons_layout.addStretch()
        buttons_layout.addWidget(self.translation_status, 1)

        layout.addLayout(buttons_layout)
        self.translation_tab.setLayout(layout)

    def _populate_translation_rows(self) -> None:
        rows = self._settings.get("translation_rows", ServiceConfig.copy_translation_rows())
        self.translation_table.setRowCount(0)
        for row in rows:
            self._append_translation_row(row)

    def _append_translation_row(self, row: Dict[str, object]) -> None:
        row_idx = self.translation_table.rowCount()
        self.translation_table.insertRow(row_idx)

        key_item = QTableWidgetItem(str(row.get("key", "")))
        name_item = QTableWidgetItem(str(row.get("name", "")))
        key_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEditable | Qt.ItemIsEnabled)
        name_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEditable | Qt.ItemIsEnabled)

        multiplier_spin = QDoubleSpinBox()
        multiplier_spin.setDecimals(2)
        multiplier_spin.setRange(0.0, 100.0)
        multiplier_spin.setValue(float(row.get("multiplier", 0.0) or 0.0))

        base_item = QTableWidgetItem()
        base_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
        base_item.setCheckState(Qt.Checked if row.get("is_base") else Qt.Unchecked)

        self.translation_table.setItem(row_idx, 0, key_item)
        self.translation_table.setItem(row_idx, 1, name_item)
        self.translation_table.setCellWidget(row_idx, 2, multiplier_spin)
        self.translation_table.setItem(row_idx, 3, base_item)

    def _add_translation_row(self) -> None:
        self._append_translation_row({"key": "", "name": tr("Новая строка", self.lang), "multiplier": 1.0, "is_base": False})
        self.translation_status.setText(tr("Добавлена строка. Проверьте значения перед сохранением.", self.lang))

    def _remove_translation_rows(self) -> None:
        rows = sorted({index.row() for index in self.translation_table.selectedIndexes()}, reverse=True)
        if not rows:
            self.translation_status.setText(tr("Выберите строки для удаления.", self.lang))
            return
        if self.translation_table.rowCount() - len(rows) < 1:
            self.translation_status.setText(tr("Должна остаться хотя бы одна строка.", self.lang))
            return
        for row in rows:
            self.translation_table.removeRow(row)
        self.translation_status.setText(tr("Строки удалены. Сохраните изменения.", self.lang))

    def _collect_translation_rows(self) -> List[Dict[str, object]]:
        rows: List[Dict[str, object]] = []
        base_seen = False
        for row in range(self.translation_table.rowCount()):
            key_item = self.translation_table.item(row, 0)
            name_item = self.translation_table.item(row, 1)
            base_item = self.translation_table.item(row, 3)
            spin = self.translation_table.cellWidget(row, 2)
            multiplier = float(spin.value()) if isinstance(spin, QDoubleSpinBox) else 0.0
            is_base = base_item.checkState() == Qt.Checked if base_item else False
            if is_base:
                if base_seen:
                    is_base = False
                    if base_item:
                        base_item.setCheckState(Qt.Unchecked)
                else:
                    base_seen = True
            rows.append(
                {
                    "key": key_item.text().strip() if key_item else "",
                    "name": name_item.text().strip() if name_item else "",
                    "multiplier": multiplier,
                    "is_base": is_base,
                }
            )
        if not base_seen and rows:
            rows[0]["is_base"] = True
        return rows

    # Additional services tab -------------------------------------------
    def _setup_additional_tab(self) -> None:
        layout = QVBoxLayout()
        self.additional_tree = QTreeWidget()
        self.additional_tree.setColumnCount(3)
        self.additional_tree.setHeaderLabels(
            [
                tr("Категория / Услуга", self.lang),
                tr("Коэффициент", self.lang),
                tr("Базовая", self.lang),
            ]
        )
        self.additional_tree.setIndentation(20)
        self.additional_tree.setItemsExpandable(True)
        self.additional_tree.setExpandsOnDoubleClick(True)
        self.additional_tree.setSelectionMode(QAbstractItemView.ExtendedSelection)
        layout.addWidget(self.additional_tree)

        buttons_layout = QHBoxLayout()
        self.add_category_btn = QPushButton(tr("Добавить категорию", self.lang))
        self.add_service_btn = QPushButton(tr("Добавить услугу", self.lang))
        self.remove_service_btn = QPushButton(tr("Удалить", self.lang))
        self.additional_status = QLabel()
        self.additional_status.setWordWrap(True)

        self.add_category_btn.clicked.connect(self._add_category)
        self.add_service_btn.clicked.connect(self._add_service)
        self.remove_service_btn.clicked.connect(self._remove_services)

        buttons_layout.addWidget(self.add_category_btn)
        buttons_layout.addWidget(self.add_service_btn)
        buttons_layout.addWidget(self.remove_service_btn)
        buttons_layout.addStretch()
        buttons_layout.addWidget(self.additional_status, 1)

        layout.addLayout(buttons_layout)
        self.additional_tab.setLayout(layout)

    def _populate_additional_services(self) -> None:
        self.additional_tree.clear()
        data = self._settings.get("additional_services", ServiceConfig.copy_additional_services())
        for section, services in data.items():
            parent = QTreeWidgetItem([section])
            parent.setFlags(Qt.ItemIsEditable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.additional_tree.addTopLevelItem(parent)
            for row in services:
                self._append_service(parent, row)
            self.additional_tree.expandItem(parent)

    def _append_service(self, parent: QTreeWidgetItem, row: Dict[str, object]) -> None:
        child = QTreeWidgetItem(parent, [str(row.get("name", ""))])
        child.setFlags(Qt.ItemIsEditable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        spin = QDoubleSpinBox()
        spin.setRange(0.0, 100.0)
        spin.setDecimals(2)
        spin.setValue(float(row.get("multiplier", 0.0) or 0.0))
        self.additional_tree.setItemWidget(child, 1, spin)
        base_item = QCheckBox()
        base_item.setChecked(bool(row.get("is_base", False)))
        base_item.setTristate(False)
        self.additional_tree.setItemWidget(child, 2, base_item)

    def _add_category(self) -> None:
        new_item = QTreeWidgetItem([tr("Новая категория", self.lang)])
        new_item.setFlags(Qt.ItemIsEditable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        self.additional_tree.addTopLevelItem(new_item)
        self.additional_tree.editItem(new_item)
        self.additional_status.setText(tr("Категория добавлена.", self.lang))

    def _add_service(self) -> None:
        selected = self.additional_tree.selectedItems()
        parent = selected[0] if selected else None
        if parent and parent.parent() is not None:
            parent = parent.parent()
        if parent is None:
            parent = self.additional_tree.topLevelItem(0)
            if parent is None:
                self._add_category()
                parent = self.additional_tree.topLevelItem(0)
        self._append_service(parent, {"name": tr("Новая услуга", self.lang), "multiplier": 1.0, "is_base": True})
        self.additional_status.setText(tr("Услуга добавлена.", self.lang))

    def _remove_services(self) -> None:
        selected = self.additional_tree.selectedItems()
        if not selected:
            self.additional_status.setText(tr("Выберите элементы для удаления.", self.lang))
            return
        for item in selected:
            parent = item.parent()
            if parent:
                parent.removeChild(item)
            else:
                index = self.additional_tree.indexOfTopLevelItem(item)
                if index >= 0:
                    self.additional_tree.takeTopLevelItem(index)
        self.additional_status.setText(tr("Элементы удалены.", self.lang))

    def _collect_additional_services(self) -> Dict[str, List[Dict[str, object]]]:
        data: Dict[str, List[Dict[str, object]]] = {}
        for idx in range(self.additional_tree.topLevelItemCount()):
            parent = self.additional_tree.topLevelItem(idx)
            section = parent.text(0).strip()
            if not section:
                continue
            rows: List[Dict[str, object]] = []
            for child_index in range(parent.childCount()):
                child = parent.child(child_index)
                spin_widget = self.additional_tree.itemWidget(child, 1)
                checkbox_widget = self.additional_tree.itemWidget(child, 2)
                multiplier = float(spin_widget.value()) if isinstance(spin_widget, QDoubleSpinBox) else 0.0
                is_base = bool(checkbox_widget.isChecked()) if isinstance(checkbox_widget, QCheckBox) else False
                rows.append(
                    {
                        "name": child.text(0).strip(),
                        "multiplier": multiplier,
                        "is_base": is_base,
                    }
                )
            if rows:
                data[section] = rows
        return data

    # Legal entities tab -------------------------------------------------
    def _setup_legal_tab(self) -> None:
        layout = QHBoxLayout()

        self.legal_list = QListWidget()
        self.legal_list.currentItemChanged.connect(self._on_legal_selection_changed)
        layout.addWidget(self.legal_list, 1)

        right = QVBoxLayout()

        form = QGridLayout()
        self.legal_name_label = QLabel()
        self.legal_template_edit = QTextEdit()
        self.legal_template_edit.setReadOnly(True)
        self.legal_template_edit.setFixedHeight(80)
        self.legal_vat_checkbox = QCheckBox(tr("VAT включён", self.lang))
        self.legal_vat_spin = QDoubleSpinBox()
        self.legal_vat_spin.setRange(0.0, 100.0)
        self.legal_vat_spin.setDecimals(2)

        form.addWidget(QLabel(tr("Название", self.lang)), 0, 0)
        form.addWidget(self.legal_name_label, 0, 1)
        form.addWidget(QLabel(tr("Шаблон", self.lang)), 1, 0)
        form.addWidget(self.legal_template_edit, 1, 1)
        form.addWidget(self.legal_vat_checkbox, 2, 0, 1, 2)
        form.addWidget(QLabel(tr("НДС по умолчанию", self.lang)), 3, 0)
        form.addWidget(self.legal_vat_spin, 3, 1)

        right.addLayout(form)

        button_row = QHBoxLayout()
        self.add_legal_btn = QPushButton(tr("Добавить шаблон", self.lang))
        self.replace_legal_btn = QPushButton(tr("Заменить файл", self.lang))
        self.remove_legal_btn = QPushButton(tr("Удалить", self.lang))
        self.export_legal_btn = QPushButton(tr("Экспортировать", self.lang))

        self.add_legal_btn.clicked.connect(self._on_add_legal_entity)
        self.replace_legal_btn.clicked.connect(self._on_replace_template)
        self.remove_legal_btn.clicked.connect(self._on_remove_legal_entity)
        self.export_legal_btn.clicked.connect(self._on_export_templates)

        button_row.addWidget(self.add_legal_btn)
        button_row.addWidget(self.replace_legal_btn)
        button_row.addWidget(self.remove_legal_btn)
        button_row.addStretch()
        button_row.addWidget(self.export_legal_btn)

        right.addLayout(button_row)

        self.legal_status = QLabel()
        self.legal_status.setWordWrap(True)
        right.addWidget(self.legal_status)
        right.addStretch()

        self.legal_vat_checkbox.stateChanged.connect(self._on_legal_metadata_changed)
        self.legal_vat_spin.valueChanged.connect(self._on_legal_metadata_changed)

        layout.addLayout(right, 2)
        self.legal_tab.setLayout(layout)

    def _populate_legal_entities(self) -> None:
        self.legal_list.clear()
        for name in sorted(self._staged_entities):
            self._add_legal_list_item(name)
        if self.legal_list.count():
            self.legal_list.setCurrentRow(0)

    def _add_legal_list_item(self, name: str) -> None:
        entity = self._staged_entities[name]
        suffix = tr(" (встроено)", self.lang) if entity.source == "built-in" else ""
        item = QListWidgetItem(f"{name}{suffix}")
        item.setData(Qt.UserRole, name)
        self.legal_list.addItem(item)

    def _current_entity(self) -> Optional[_StagedEntity]:
        item = self.legal_list.currentItem()
        if not item:
            return None
        name = item.data(Qt.UserRole)
        if not name:
            return None
        return self._staged_entities.get(name)

    def _on_legal_selection_changed(self, current: QListWidgetItem | None, _previous: QListWidgetItem | None) -> None:
        entity = self._current_entity()
        if not entity:
            self.legal_name_label.setText("")
            self.legal_template_edit.setPlainText("")
            self.legal_vat_checkbox.setEnabled(False)
            self.legal_vat_spin.setEnabled(False)
            return
        self.legal_name_label.setText(entity.name)
        self.legal_template_edit.setPlainText(entity.effective_template())
        vat_enabled = bool(entity.metadata.get("vat_enabled", False))
        vat_default = float(entity.metadata.get("default_vat", 0.0) or 0.0)
        self.legal_vat_checkbox.blockSignals(True)
        self.legal_vat_spin.blockSignals(True)
        self.legal_vat_checkbox.setChecked(vat_enabled)
        self.legal_vat_spin.setValue(vat_default)
        self.legal_vat_checkbox.blockSignals(False)
        self.legal_vat_spin.blockSignals(False)
        editable = entity.source != "built-in"
        self.replace_legal_btn.setEnabled(editable)
        self.remove_legal_btn.setEnabled(editable)
        self.legal_vat_checkbox.setEnabled(editable)
        self.legal_vat_spin.setEnabled(editable)

    def _on_legal_metadata_changed(self) -> None:
        entity = self._current_entity()
        if not entity or entity.source == "built-in":
            return
        entity.metadata["vat_enabled"] = bool(self.legal_vat_checkbox.isChecked())
        entity.metadata["default_vat"] = float(self.legal_vat_spin.value())
        self.legal_status.setText(tr("Параметры обновлены. Не забудьте сохранить.", self.lang))

    def _on_add_legal_entity(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(self, tr("Выберите шаблон", self.lang), str(Path.home()), "Excel (*.xlsx)")
        if not file_path:
            return
        name, ok = QInputDialog.getText(self, tr("Название юрлица", self.lang), tr("Введите название", self.lang))
        if not ok or not name.strip():
            QMessageBox.warning(self, tr("Ошибка", self.lang), tr("Название не может быть пустым.", self.lang))
            return
        name = name.strip()
        if name in self._staged_entities and not self._staged_entities[name].to_delete:
            QMessageBox.warning(self, tr("Ошибка", self.lang), tr("Юрлицо с таким названием уже существует.", self.lang))
            return
        entity = _StagedEntity(
            name=name,
            source="user",
            original_template="",
            pending_template=file_path,
            metadata={"vat_enabled": False, "default_vat": 0.0},
            is_new=True,
        )
        self._staged_entities[name] = entity
        self._add_legal_list_item(name)
        self.legal_list.setCurrentRow(self.legal_list.count() - 1)
        self.legal_status.setText(tr("Шаблон добавлен. Сохраните изменения для применения.", self.lang))

    def _on_replace_template(self) -> None:
        entity = self._current_entity()
        if not entity or entity.source == "built-in":
            return
        file_path, _ = QFileDialog.getOpenFileName(self, tr("Выберите шаблон", self.lang), str(Path.home()), "Excel (*.xlsx)")
        if not file_path:
            return
        entity.pending_template = file_path
        self.legal_template_edit.setPlainText(entity.effective_template())
        self.legal_status.setText(tr("Файл обновлён. Сохраните изменения для применения.", self.lang))

    def _on_remove_legal_entity(self) -> None:
        entity = self._current_entity()
        if not entity or entity.source == "built-in":
            return
        reply = QMessageBox.question(
            self,
            tr("Подтверждение", self.lang),
            tr("Удалить выбранный шаблон?", self.lang),
        )
        if reply != QMessageBox.Yes:
            return
        entity.to_delete = True
        row = self.legal_list.currentRow()
        if row >= 0:
            self.legal_list.takeItem(row)
        self.legal_status.setText(tr("Шаблон помечен на удаление. Сохраните изменения.", self.lang))

    def _on_export_templates(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            tr("Экспорт шаблонов", self.lang),
            str(Path.home() / "legal_entities.xlsx"),
            "Excel (*.xlsx)",
        )
        if not path:
            return
        try:
            records = list(self._export_records_view())
            export_legal_entities_to_excel(Path(path), records)
            self.legal_status.setText(tr("Экспорт успешно завершён.", self.lang))
        except Exception as exc:  # pragma: no cover - UI feedback
            QMessageBox.warning(
                self,
                tr("Ошибка", self.lang),
                tr("Не удалось экспортировать шаблоны: {}", self.lang).format(exc),
            )

    def _export_records_view(self) -> Iterable[LegalEntityRecord]:
        seen: set[str] = set()
        for name, record in self._initial_records.items():
            staged = self._staged_entities.get(name)
            if staged and staged.to_delete:
                continue
            if staged:
                template = staged.effective_template() or record.template
                metadata = dict(staged.metadata)
                source = staged.source
            else:
                template = record.template
                metadata = dict(record.metadata or {})
                source = record.source
            seen.add(name)
            yield LegalEntityRecord(name=name, template=template, metadata=metadata, source=source)
        for name, staged in self._staged_entities.items():
            if name in seen or staged.to_delete:
                continue
            template = staged.effective_template()
            yield LegalEntityRecord(name=name, template=template, metadata=dict(staged.metadata), source=staged.source)

    # Collect and save ---------------------------------------------------
    def _populate_fuzzy_thresholds(self) -> None:
        thresholds = self._settings.get("fuzzy_thresholds", ServiceConfig.FUZZY_THRESHOLDS)
        self.new_words_spin.setValue(int(thresholds.get("new_words", 100)))
        self.other_spin.setValue(int(thresholds.get("others", 75)))

    def _gather_settings(self) -> Dict[str, object]:
        data = {
            "translation_rows": self._collect_translation_rows(),
            "additional_services": self._collect_additional_services(),
            "fuzzy_thresholds": {
                "new_words": int(self.new_words_spin.value()),
                "others": int(self.other_spin.value()),
            },
        }
        return data

    def _apply_legal_changes(self) -> None:
        # Remove entities marked for deletion first
        for entity in list(self._staged_entities.values()):
            if entity.to_delete and entity.source != "built-in":
                remove_user_legal_entity(entity.name)

        # Apply additions and updates
        for entity in self._staged_entities.values():
            if entity.to_delete:
                continue
            if entity.source == "built-in" and not entity.is_new:
                continue
            template_path = entity.pending_template or entity.original_template
            if not template_path:
                continue
            try:
                add_or_update_legal_entity(entity.name, Path(template_path), dict(entity.metadata))
            except Exception as exc:  # pragma: no cover - UI feedback
                QMessageBox.warning(
                    self,
                    tr("Ошибка", self.lang),
                    tr("Не удалось сохранить шаблон {}: {}", self.lang).format(entity.name, exc),
                )

    def _on_accept(self) -> None:
        settings = self._gather_settings()
        try:
            save_settings(settings)
            self._apply_legal_changes()
        except Exception as exc:  # pragma: no cover - UI feedback
            QMessageBox.warning(
                self,
                tr("Ошибка", self.lang),
                tr("Не удалось сохранить настройки: {}", self.lang).format(exc),
            )
            return
        self.accept()
