"""Settings dialog that allows managing templates and service coefficients."""

from __future__ import annotations

from dataclasses import dataclass, field
import shutil
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
    QLineEdit,
    QHeaderView,
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
    pending_logo: Optional[str] = None
    is_new: bool = False
    to_delete: bool = False

    def effective_template(self) -> str:
        return self.pending_template or self.original_template

    def effective_logo(self) -> Optional[str]:
        if self.pending_logo is not None:
            return self.pending_logo or None
        logo = self.metadata.get("logo") if isinstance(self.metadata, dict) else None
        return str(logo) if logo else None


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
        self._translation_keys: List[str] = []

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

        self.translation_table = QTableWidget(0, 2)
        self.translation_table.setHorizontalHeaderLabels(
            [
                tr("Название", self.lang),
                tr("Процент", self.lang),
            ]
        )
        header = self.translation_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.translation_table.verticalHeader().setVisible(False)
        self.translation_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.translation_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.translation_table.setEditTriggers(
            QAbstractItemView.DoubleClicked
            | QAbstractItemView.EditKeyPressed
            | QAbstractItemView.SelectedClicked
        )

        layout.addWidget(self.translation_table)

        self.translation_status = QLabel()
        self.translation_status.setWordWrap(True)
        layout.addWidget(self.translation_status)
        self.translation_tab.setLayout(layout)

    def _populate_translation_rows(self) -> None:
        rows = self._settings.get("translation_rows", ServiceConfig.copy_translation_rows())
        defaults = ServiceConfig.copy_translation_rows()
        prepared: List[Dict[str, object]] = []
        seen_keys: set[str] = set()
        target_rows = 4
        for row in rows:
            if len(prepared) >= target_rows or not isinstance(row, dict):
                continue
            key = str(row.get("key", ""))
            if key and key in seen_keys:
                continue
            prepared.append(dict(row))
            if key:
                seen_keys.add(key)
        for row in defaults:
            if len(prepared) >= target_rows:
                break
            key = str(row.get("key", ""))
            if key and key in seen_keys:
                continue
            prepared.append(dict(row))
            if key:
                seen_keys.add(key)
        while len(prepared) < target_rows:
            key = f"custom_{len(prepared)}"
            prepared.append(
                {
                    "key": key,
                    "name": tr("Новая строка", self.lang),
                    "multiplier": 1.0,
                    "is_base": False,
                }
            )

        self.translation_table.blockSignals(True)
        self.translation_table.setRowCount(len(prepared))
        self._translation_keys = []
        for row_idx, row in enumerate(prepared):
            key = str(row.get("key") or f"row_{row_idx}")
            if key in self._translation_keys:
                key = f"{key}_{row_idx}"
            self._translation_keys.append(key)

            name_item = QTableWidgetItem(str(row.get("name", "")))
            name_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEditable | Qt.ItemIsEnabled)
            self.translation_table.setItem(row_idx, 0, name_item)

            spin = QDoubleSpinBox()
            spin.setDecimals(2)
            spin.setRange(0.0, 100.0)
            spin.setSuffix("%")
            spin.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            raw_multiplier = float(row.get("multiplier", 0.0) or 0.0)
            percent = raw_multiplier if raw_multiplier > 1.0 else raw_multiplier * 100.0
            spin.setValue(percent)
            self.translation_table.setCellWidget(row_idx, 1, spin)
        self.translation_table.blockSignals(False)
        self.translation_status.clear()

    def _collect_translation_rows(self) -> List[Dict[str, object]]:
        rows: List[Dict[str, object]] = []
        for row in range(self.translation_table.rowCount()):
            name_item = self.translation_table.item(row, 0)
            spin = self.translation_table.cellWidget(row, 1)
            percent = float(spin.value()) if isinstance(spin, QDoubleSpinBox) else 0.0
            multiplier = percent / 100.0
            key = self._translation_keys[row] if row < len(self._translation_keys) else f"row_{row}"
            rows.append(
                {
                    "key": key,
                    "name": name_item.text().strip() if name_item else "",
                    "multiplier": multiplier,
                    "is_base": row == 0,
                }
            )
        if rows and not rows[0]["name"]:
            rows[0]["name"] = tr("Перевод, новые слова", self.lang)
        return rows

    # Additional services tab -------------------------------------------
    def _setup_additional_tab(self) -> None:
        layout = QVBoxLayout()
        self.additional_list = QListWidget()
        self.additional_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.additional_list.setEditTriggers(
            QAbstractItemView.DoubleClicked
            | QAbstractItemView.EditKeyPressed
            | QAbstractItemView.SelectedClicked
        )
        layout.addWidget(self.additional_list)

        buttons_layout = QHBoxLayout()
        self.add_service_btn = QPushButton(tr("Добавить услугу", self.lang))
        self.remove_service_btn = QPushButton(tr("Удалить", self.lang))
        self.additional_status = QLabel()
        self.additional_status.setWordWrap(True)

        self.add_service_btn.clicked.connect(self._add_service_name)
        self.remove_service_btn.clicked.connect(self._remove_service_names)

        buttons_layout.addWidget(self.add_service_btn)
        buttons_layout.addWidget(self.remove_service_btn)
        buttons_layout.addStretch()
        buttons_layout.addWidget(self.additional_status, 1)

        layout.addLayout(buttons_layout)
        self.additional_tab.setLayout(layout)

    def _populate_additional_services(self) -> None:
        self.additional_list.clear()
        data = self._settings.get("additional_services", ServiceConfig.copy_additional_services())
        if not isinstance(data, list):
            data = list(data) if data else []
        if not data:
            data = []
        for name in data:
            item = QListWidgetItem(str(name))
            item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEditable | Qt.ItemIsEnabled)
            self.additional_list.addItem(item)
        if self.additional_list.count() == 0:
            self._add_service_name(edit_immediately=False)

    def _add_service_name(self, edit_immediately: bool = True) -> None:
        item = QListWidgetItem(tr("Новая услуга", self.lang))
        item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEditable | Qt.ItemIsEnabled)
        self.additional_list.addItem(item)
        self.additional_list.setCurrentItem(item)
        if edit_immediately:
            self.additional_list.editItem(item)
        self.additional_status.setText(tr("Услуга добавлена.", self.lang))

    def _remove_service_names(self) -> None:
        selected = self.additional_list.selectedItems()
        if not selected:
            self.additional_status.setText(tr("Выберите элементы для удаления.", self.lang))
            return
        for item in selected:
            row = self.additional_list.row(item)
            self.additional_list.takeItem(row)
        self.additional_status.setText(tr("Элементы удалены.", self.lang))

    def _collect_additional_services(self) -> List[str]:
        services: List[str] = []
        for idx in range(self.additional_list.count()):
            text = self.additional_list.item(idx).text().strip()
            if text:
                services.append(text)
        return services

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
        self.legal_logo_edit = QLineEdit()
        self.legal_logo_edit.setReadOnly(True)

        form.addWidget(QLabel(tr("Название", self.lang)), 0, 0)
        form.addWidget(self.legal_name_label, 0, 1)
        form.addWidget(QLabel(tr("Шаблон", self.lang)), 1, 0)
        form.addWidget(self.legal_template_edit, 1, 1)
        form.addWidget(self.legal_vat_checkbox, 2, 0, 1, 2)
        form.addWidget(QLabel(tr("НДС по умолчанию", self.lang)), 3, 0)
        form.addWidget(self.legal_vat_spin, 3, 1)
        form.addWidget(QLabel(tr("Логотип", self.lang)), 4, 0)

        logo_row = QHBoxLayout()
        logo_row.setContentsMargins(0, 0, 0, 0)
        logo_row.addWidget(self.legal_logo_edit)
        self.choose_logo_btn = QPushButton(tr("Выбрать логотип", self.lang))
        self.clear_logo_btn = QPushButton(tr("Удалить логотип", self.lang))
        logo_row.addWidget(self.choose_logo_btn)
        logo_row.addWidget(self.clear_logo_btn)
        form.addLayout(logo_row, 4, 1)

        right.addLayout(form)

        button_row = QHBoxLayout()
        self.download_template_btn = QPushButton(tr("Скачать шаблон", self.lang))
        self.add_legal_btn = QPushButton(tr("Добавить шаблон", self.lang))
        self.replace_legal_btn = QPushButton(tr("Заменить файл", self.lang))
        self.remove_legal_btn = QPushButton(tr("Удалить", self.lang))
        self.export_legal_btn = QPushButton(tr("Экспортировать", self.lang))

        self.download_template_btn.clicked.connect(self._on_download_template)
        self.add_legal_btn.clicked.connect(self._on_add_legal_entity)
        self.replace_legal_btn.clicked.connect(self._on_replace_template)
        self.remove_legal_btn.clicked.connect(self._on_remove_legal_entity)
        self.export_legal_btn.clicked.connect(self._on_export_templates)
        self.choose_logo_btn.clicked.connect(self._on_choose_logo)
        self.clear_logo_btn.clicked.connect(self._on_clear_logo)

        button_row.addWidget(self.download_template_btn)
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
            self.legal_logo_edit.clear()
            self.choose_logo_btn.setEnabled(False)
            self.clear_logo_btn.setEnabled(False)
            self.download_template_btn.setEnabled(False)
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
        logo_path = entity.effective_logo() or ""
        self.legal_logo_edit.setText(logo_path)
        self.choose_logo_btn.setEnabled(editable)
        self.clear_logo_btn.setEnabled(editable and bool(logo_path))
        self.download_template_btn.setEnabled(bool(entity.effective_template()))

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
            pending_logo=None,
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
        self.download_template_btn.setEnabled(True)
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

    def _on_download_template(self) -> None:
        entity = self._current_entity()
        if not entity:
            return
        template_path = entity.effective_template()
        if not template_path:
            QMessageBox.warning(self, tr("Ошибка", self.lang), tr("Шаблон недоступен для скачивания.", self.lang))
            return
        source = Path(template_path)
        default_name = source.name if source.name else f"{entity.name}.xlsx"
        target, _ = QFileDialog.getSaveFileName(
            self,
            tr("Сохранить шаблон", self.lang),
            str(Path.home() / default_name),
            "Excel (*.xlsx)",
        )
        if not target:
            return
        try:
            shutil.copy2(template_path, target)
            self.legal_status.setText(tr("Шаблон сохранён: {}", self.lang).format(Path(target).name))
        except Exception as exc:  # pragma: no cover - filesystem dependent
            QMessageBox.warning(
                self,
                tr("Ошибка", self.lang),
                tr("Не удалось сохранить шаблон: {}", self.lang).format(exc),
            )

    def _on_choose_logo(self) -> None:
        entity = self._current_entity()
        if not entity or entity.source == "built-in":
            return
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            tr("Выберите логотип", self.lang),
            str(Path.home()),
            "Images (*.png *.jpg *.jpeg *.bmp)",
        )
        if not file_path:
            return
        entity.pending_logo = file_path
        self.legal_logo_edit.setText(file_path)
        self.clear_logo_btn.setEnabled(True)
        self.legal_status.setText(tr("Логотип будет обновлён после сохранения.", self.lang))

    def _on_clear_logo(self) -> None:
        entity = self._current_entity()
        if not entity or entity.source == "built-in":
            return
        entity.pending_logo = ""
        self.legal_logo_edit.clear()
        self.clear_logo_btn.setEnabled(False)
        self.legal_status.setText(tr("Логотип будет удалён после сохранения.", self.lang))

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
                if staged.pending_logo == "":
                    metadata.pop("logo", None)
                elif staged.pending_logo:
                    metadata["logo"] = staged.pending_logo
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
            metadata = dict(staged.metadata)
            if staged.pending_logo == "":
                metadata.pop("logo", None)
            elif staged.pending_logo:
                metadata["logo"] = staged.pending_logo
            yield LegalEntityRecord(name=name, template=template, metadata=metadata, source=staged.source)

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
            metadata = dict(entity.metadata)
            if entity.pending_logo == "":
                metadata.pop("logo", None)
            elif entity.pending_logo:
                metadata["logo"] = entity.pending_logo
            try:
                add_or_update_legal_entity(entity.name, Path(template_path), metadata)
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
