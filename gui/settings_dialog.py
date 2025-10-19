"""Settings dialog that allows managing legal entity templates."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from gui.logo_label import ScaledPixmapLabel

from logic.legal_entities import (
    LegalEntityRecord,
    add_or_update_legal_entity,
    add_or_update_logo,
    get_builtin_logo_path,
    export_builtin_templates,
    export_legal_entities_to_excel,
    get_logo_path,
    get_logo_source,
    list_legal_entities_with_sources,
    remove_user_legal_entity,
    remove_logo_override,
)
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
    logo_path: Optional[str] = None
    logo_source: str = ""
    pending_logo: Optional[str] = None
    remove_logo: bool = False
    builtin_logo_path: Optional[str] = None

    def effective_template(self) -> str:
        return self.pending_template or self.original_template

    def effective_logo(self) -> Optional[str]:
        if self.pending_logo:
            return self.pending_logo
        if self.remove_logo:
            return self.builtin_logo_path
        if self.logo_path:
            return self.logo_path
        return self.builtin_logo_path


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
                logo_path=get_logo_path(name),
                logo_source=get_logo_source(name),
                builtin_logo_path=get_builtin_logo_path(name),
            )
            for name, record in self._initial_records.items()
        }

        self.setWindowTitle(tr("Настройки", self.lang))
        self.resize(960, 640)

        self._setup_ui()
        self._populate_legal_entities()

    # ------------------------------------------------------------------ UI
    def _setup_ui(self) -> None:
        layout = QVBoxLayout(self)
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)

        self.legal_tab = QWidget()
        self._setup_legal_tab()
        self.tabs.addTab(self.legal_tab, tr("Юрлица", self.lang))

        self.buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        self.buttons.accepted.connect(self._on_accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)

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
        self.download_builtin_btn = QPushButton(tr("Скачать встроенные", self.lang))
        self.export_legal_btn = QPushButton(tr("Экспортировать", self.lang))

        self.add_legal_btn.clicked.connect(self._on_add_legal_entity)
        self.replace_legal_btn.clicked.connect(self._on_replace_template)
        self.remove_legal_btn.clicked.connect(self._on_remove_legal_entity)
        self.download_builtin_btn.clicked.connect(self._on_download_builtin_templates)
        self.export_legal_btn.clicked.connect(self._on_export_templates)

        button_row.addWidget(self.add_legal_btn)
        button_row.addWidget(self.replace_legal_btn)
        button_row.addWidget(self.remove_legal_btn)
        button_row.addWidget(self.download_builtin_btn)
        button_row.addStretch()
        button_row.addWidget(self.export_legal_btn)

        right.addLayout(button_row)

        self.logo_preview = ScaledPixmapLabel()
        self.logo_preview.setMinimumHeight(140)
        self.logo_preview.setStyleSheet("border: 1px dashed #c0c0c0; background: #fafafa;")
        self.logo_preview.setText(tr("Логотип не выбран", self.lang))
        right.addWidget(self.logo_preview)

        self.logo_info_label = QLabel()
        self.logo_info_label.setWordWrap(True)
        right.addWidget(self.logo_info_label)

        logo_buttons = QHBoxLayout()
        self.change_logo_btn = QPushButton(tr("Изменить логотип", self.lang))
        self.remove_logo_btn = QPushButton(tr("Удалить логотип", self.lang))
        self.change_logo_btn.clicked.connect(self._on_change_logo)
        self.remove_logo_btn.clicked.connect(self._on_remove_logo)
        logo_buttons.addWidget(self.change_logo_btn)
        logo_buttons.addWidget(self.remove_logo_btn)
        logo_buttons.addStretch()
        right.addLayout(logo_buttons)

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
            self.change_logo_btn.setEnabled(False)
            self.remove_logo_btn.setEnabled(False)
            self._update_logo_preview(None)
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
        self.change_logo_btn.setEnabled(True)
        self._update_logo_preview(entity)

    def _on_legal_metadata_changed(self) -> None:
        entity = self._current_entity()
        if not entity or entity.source == "built-in":
            return
        entity.metadata["vat_enabled"] = bool(self.legal_vat_checkbox.isChecked())
        entity.metadata["default_vat"] = float(self.legal_vat_spin.value())
        self.legal_status.setText(tr("Параметры обновлены. Не забудьте сохранить.", self.lang))

    def _on_change_logo(self) -> None:
        entity = self._current_entity()
        if not entity:
            return
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            tr("Выберите логотип", self.lang),
            str(Path.home()),
            "Images (*.png *.jpg *.jpeg)",
        )
        if not file_path:
            return
        entity.pending_logo = file_path
        entity.remove_logo = False
        self.legal_status.setText(tr("Логотип обновлён. Сохраните изменения для применения.", self.lang))
        self._update_logo_preview(entity)

    def _on_remove_logo(self) -> None:
        entity = self._current_entity()
        if not entity:
            return
        entity.pending_logo = None
        entity.remove_logo = True
        self.legal_status.setText(tr("Пользовательский логотип будет удалён после сохранения.", self.lang))
        self._update_logo_preview(entity)

    def _update_logo_preview(self, entity: Optional[_StagedEntity]) -> None:
        if not entity:
            self.logo_preview.clear()
            self.logo_preview.setText(tr("Логотип не выбран", self.lang))
            self.logo_info_label.setText("")
            self.logo_preview.setToolTip("")
            return

        effective_logo = entity.effective_logo()
        info_parts: List[str] = []
        if entity.pending_logo:
            info_parts.append(tr("Новый логотип (не сохранён)", self.lang))
        elif entity.remove_logo:
            info_parts.append(tr("Логотип будет удалён", self.lang))
        else:
            if entity.logo_source == "user":
                info_parts.append(tr("Источник: пользовательский", self.lang))
            elif entity.logo_source == "built-in":
                info_parts.append(tr("Источник: встроенный", self.lang))

        if effective_logo and self.logo_preview.set_path(effective_logo):
            self.logo_info_label.setText("\n".join(info_parts + [effective_logo]))
            self.logo_preview.setToolTip(effective_logo)
        else:
            if entity.remove_logo and entity.builtin_logo_path:
                # Pending removal but built-in logo remains available
                self.logo_preview.set_path(entity.builtin_logo_path)
                info_parts.append(entity.builtin_logo_path)
                self.logo_info_label.setText("\n".join(info_parts))
                self.logo_preview.setToolTip(entity.builtin_logo_path or "")
            else:
                self.logo_preview.clear()
                self.logo_preview.setText(tr("Логотип отсутствует", self.lang))
                self.logo_info_label.setText("\n".join(info_parts))
                self.logo_preview.setToolTip("")

        has_override = bool(entity.pending_logo or (entity.logo_source == "user" and not entity.remove_logo))
        self.remove_logo_btn.setEnabled(has_override)

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
            logo_path=None,
            logo_source="user",
            builtin_logo_path=get_builtin_logo_path(name),
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

    def _on_download_builtin_templates(self) -> None:
        target_dir = QFileDialog.getExistingDirectory(
            self,
            tr("Выберите папку для сохранения", self.lang),
            str(Path.home()),
        )
        if not target_dir:
            return
        try:
            exported = export_builtin_templates(Path(target_dir))
        except Exception as exc:  # pragma: no cover - UI feedback
            QMessageBox.warning(
                self,
                tr("Ошибка", self.lang),
                tr("Не удалось сохранить встроенные шаблоны: {}", self.lang).format(exc),
            )
            return
        if not exported:
            self.legal_status.setText(tr("Встроенные шаблоны не найдены.", self.lang))
        else:
            self.legal_status.setText(
                tr("Сохранено встроенных шаблонов: {}", self.lang).format(len(exported))
            )

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
    def _gather_settings(self) -> Dict[str, object]:
        """Return settings unchanged to preserve existing configuration."""

        return dict(self._settings)

    def _apply_legal_changes(self) -> None:
        # Remove entities marked for deletion first
        for entity in list(self._staged_entities.values()):
            if entity.to_delete and entity.source != "built-in":
                remove_user_legal_entity(entity.name)
                remove_logo_override(entity.name)

        # Apply additions and updates
        for entity in self._staged_entities.values():
            if entity.to_delete:
                continue
            template_path = entity.pending_template or entity.original_template
            if entity.source != "built-in" or entity.is_new:
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

            if entity.pending_logo:
                try:
                    stored = add_or_update_logo(entity.name, Path(entity.pending_logo))
                    entity.logo_path = str(stored)
                    entity.logo_source = "user"
                    entity.builtin_logo_path = entity.builtin_logo_path or get_builtin_logo_path(entity.name)
                    entity.pending_logo = None
                    entity.remove_logo = False
                except Exception as exc:  # pragma: no cover - UI feedback
                    QMessageBox.warning(
                        self,
                        tr("Ошибка", self.lang),
                        tr("Не удалось сохранить логотип {}: {}", self.lang).format(entity.name, exc),
                    )
            elif entity.remove_logo:
                try:
                    remove_logo_override(entity.name)
                    entity.logo_path = entity.builtin_logo_path
                    entity.logo_source = "built-in" if entity.builtin_logo_path else ""
                    entity.pending_logo = None
                    entity.remove_logo = False
                except Exception as exc:  # pragma: no cover - UI feedback
                    QMessageBox.warning(
                        self,
                        tr("Ошибка", self.lang),
                        tr("Не удалось удалить логотип {}: {}", self.lang).format(entity.name, exc),
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
