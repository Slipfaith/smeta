import os
import shutil
from pathlib import Path
from typing import Dict, Optional

from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon, QPixmap
from PySide6.QtWidgets import (
    QDialog,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
    QDialogButtonBox,
    QFormLayout,
    QLineEdit,
    QCheckBox,
    QDoubleSpinBox,
)

from logic.legal_entities import (
    get_legal_entity_metadata,
    get_user_logos_dir,
    get_user_templates_dir,
    is_default_entity,
    load_legal_entities,
    load_user_entities,
    remove_user_entity,
    set_user_entity,
)
from logic.translation_config import tr


def _safe_filename(name: str) -> str:
    cleaned = "".join(ch if ch.isalnum() or ch in (" ", "-", "_", ".") else "_" for ch in name)
    cleaned = cleaned.strip().replace(" ", "_")
    return cleaned or "entity"


class LegalEntityEditor(QDialog):
    """Dialog to collect information for a new legal entity."""

    def __init__(self, lang: str = "ru", parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.lang = lang
        self._data: Optional[Dict[str, object]] = None
        self.setWindowTitle(tr("Новое юрлицо", lang))

        layout = QVBoxLayout(self)

        form = QFormLayout()
        self.name_edit = QLineEdit()
        form.addRow(tr("Название", lang) + ":", self.name_edit)

        self.template_edit = QLineEdit()
        self.template_edit.setReadOnly(True)
        template_btn = QPushButton(tr("Выбрать шаблон", lang))
        template_btn.clicked.connect(self._choose_template)
        template_row = QHBoxLayout()
        template_row.addWidget(self.template_edit)
        template_row.addWidget(template_btn)
        form.addRow(tr("Файл шаблона", lang) + ":", template_row)

        self.logo_edit = QLineEdit()
        self.logo_edit.setReadOnly(True)
        logo_btn = QPushButton(tr("Выбрать логотип", lang))
        logo_btn.clicked.connect(self._choose_logo)
        logo_row = QHBoxLayout()
        logo_row.addWidget(self.logo_edit)
        logo_row.addWidget(logo_btn)
        form.addRow(tr("Логотип", lang) + ":", logo_row)

        self.vat_checkbox = QCheckBox(tr("Включить НДС", lang))
        self.vat_checkbox.stateChanged.connect(self._toggle_vat)
        self.vat_spin = QDoubleSpinBox()
        self.vat_spin.setDecimals(2)
        self.vat_spin.setRange(0, 100)
        self.vat_spin.setValue(20.0)
        self.vat_spin.setEnabled(False)
        vat_row = QHBoxLayout()
        vat_row.addWidget(self.vat_checkbox)
        vat_row.addWidget(self.vat_spin)
        form.addRow(tr("Настройки НДС", lang) + ":", vat_row)

        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _toggle_vat(self, state: int) -> None:
        self.vat_spin.setEnabled(state == Qt.Checked)

    def _choose_template(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            tr("Выберите шаблон", self.lang),
            str(Path.home()),
            "Excel (*.xlsx *.xlsm);;All files (*.*)",
        )
        if file_path:
            self.template_edit.setText(file_path)

    def _choose_logo(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            tr("Выберите логотип", self.lang),
            str(Path.home()),
            "Images (*.png *.jpg *.jpeg *.bmp);;All files (*.*)",
        )
        if file_path:
            self.logo_edit.setText(file_path)

    def accept(self) -> None:  # type: ignore[override]
        name = self.name_edit.text().strip()
        template = self.template_edit.text().strip()
        logo = self.logo_edit.text().strip()

        if not name:
            QMessageBox.warning(
                self,
                tr("Ошибка", self.lang),
                tr("Введите название юрлица", self.lang),
            )
            return
        if not template:
            QMessageBox.warning(
                self,
                tr("Ошибка", self.lang),
                tr("Выберите файл шаблона", self.lang),
            )
            return
        if not os.path.exists(template):
            QMessageBox.warning(
                self,
                tr("Ошибка", self.lang),
                tr("Файл шаблона не найден", self.lang),
            )
            return

        data: Dict[str, object] = {
            "name": name,
            "template": template,
            "logo": logo if logo else None,
            "vat_enabled": self.vat_checkbox.isChecked(),
            "default_vat": self.vat_spin.value() if self.vat_checkbox.isChecked() else None,
        }
        self._data = data
        super().accept()

    def get_data(self) -> Optional[Dict[str, object]]:
        return self._data


class SettingsDialog(QDialog):
    """Application settings dialog for managing legal entity templates."""

    def __init__(self, lang: str = "ru", parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.lang = lang
        self._modified = False
        self.entities: Dict[str, str] = {}
        self.metadata: Dict[str, Dict[str, object]] = {}
        self.user_entities: Dict[str, Dict[str, object]] = {}

        self.setWindowTitle(tr("Настройки", lang))
        self.resize(720, 420)

        main_layout = QVBoxLayout(self)
        content_layout = QHBoxLayout()

        self.list_widget = QListWidget()
        self.list_widget.currentItemChanged.connect(self._on_selection_changed)
        content_layout.addWidget(self.list_widget, 2)

        self.details_container = QVBoxLayout()

        self.template_label = QLabel()
        self.template_label.setWordWrap(True)
        self.details_container.addWidget(self.template_label)

        self.vat_label = QLabel()
        self.details_container.addWidget(self.vat_label)

        self.source_label = QLabel()
        self.details_container.addWidget(self.source_label)

        self.logo_label = QLabel()
        self.logo_label.setFrameShape(QFrame.StyledPanel)
        self.logo_label.setAlignment(Qt.AlignCenter)
        self.logo_label.setWordWrap(True)
        self.logo_label.setMinimumSize(140, 90)
        self.logo_label.setMaximumSize(180, 120)
        self.details_container.addWidget(self.logo_label)

        button_bar = QHBoxLayout()
        self.export_btn = QPushButton(tr("Скачать шаблон", lang))
        self.export_btn.clicked.connect(self._export_template)
        button_bar.addWidget(self.export_btn)

        self.import_btn = QPushButton(tr("Импортировать шаблон", lang))
        self.import_btn.clicked.connect(self._import_template)
        button_bar.addWidget(self.import_btn)

        self.logo_btn = QPushButton(tr("Обновить логотип", lang))
        self.logo_btn.clicked.connect(self._import_logo)
        button_bar.addWidget(self.logo_btn)

        self.delete_btn = QPushButton(tr("Удалить", lang))
        self.delete_btn.clicked.connect(self._delete_entity)
        button_bar.addWidget(self.delete_btn)

        self.details_container.addLayout(button_bar)

        content_layout.addLayout(self.details_container, 3)
        main_layout.addLayout(content_layout)

        bottom_bar = QHBoxLayout()
        self.add_btn = QPushButton(tr("Добавить юрлицо", lang))
        self.add_btn.clicked.connect(self._add_entity)
        bottom_bar.addWidget(self.add_btn)
        bottom_bar.addStretch(1)

        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(self.reject)
        bottom_bar.addWidget(buttons)
        main_layout.addLayout(bottom_bar)

        self.refresh_entities()

    # ------------------------------------------------------------------
    def refresh_entities(self, selected: Optional[str] = None) -> None:
        self.entities = load_legal_entities()
        self.metadata = get_legal_entity_metadata()
        self.user_entities = load_user_entities()

        current_text = selected or self.current_entity()

        self.list_widget.blockSignals(True)
        self.list_widget.clear()
        for name in sorted(self.entities.keys()):
            item = QListWidgetItem(name)
            icon = self._make_icon(name)
            if icon is not None:
                item.setIcon(icon)
            self.list_widget.addItem(item)
            if current_text and name == current_text:
                item.setSelected(True)
        self.list_widget.blockSignals(False)

        if self.list_widget.currentItem() is None and self.list_widget.count() > 0:
            self.list_widget.setCurrentRow(0)

        self._update_details(self.current_entity())

    def was_modified(self) -> bool:
        return self._modified

    def current_entity(self) -> Optional[str]:
        item = self.list_widget.currentItem()
        return item.text() if item else None

    def _make_icon(self, entity: str) -> Optional[QIcon]:
        meta = self.metadata.get(entity, {})
        logo_path = meta.get("logo") if isinstance(meta, dict) else None
        if logo_path and isinstance(logo_path, str) and os.path.exists(logo_path):
            pixmap = QPixmap(logo_path)
            if not pixmap.isNull():
                return QIcon(pixmap)
        return None

    def _on_selection_changed(self, current: Optional[QListWidgetItem], previous) -> None:
        entity = current.text() if current else None
        self._update_details(entity)

    def _update_details(self, entity: Optional[str]) -> None:
        if not entity:
            self.template_label.setText(tr("Выберите юрлицо", self.lang))
            self.vat_label.clear()
            self.source_label.clear()
            self.logo_label.clear()
            self.delete_btn.setEnabled(False)
            return

        template_path = self.entities.get(entity, "")
        self.template_label.setText(
            tr("Шаблон", self.lang) + f":\n{template_path or tr('Не задан', self.lang)}"
        )

        meta = self.metadata.get(entity, {})
        vat_enabled = bool(meta.get("vat_enabled")) if isinstance(meta, dict) else False
        vat_value = meta.get("default_vat") if isinstance(meta, dict) else None
        if vat_enabled and vat_value is not None:
            vat_text = tr("НДС включен", self.lang) + f" ({vat_value}%)"
        elif vat_enabled:
            vat_text = tr("НДС включен", self.lang)
        else:
            vat_text = tr("НДС отключен", self.lang)
        self.vat_label.setText(vat_text)

        source = meta.get("source") if isinstance(meta, dict) else "default"
        source_text = tr("Источник", self.lang) + ": " + (
            tr("Пользователь", self.lang) if source == "user" else tr("Стандарт", self.lang)
        )
        self.source_label.setText(source_text)

        logo_path = meta.get("logo") if isinstance(meta, dict) else None
        if logo_path and isinstance(logo_path, str) and os.path.exists(logo_path):
            pixmap = QPixmap(logo_path)
            if not pixmap.isNull():
                scaled = pixmap.scaled(
                    self.logo_label.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation
                )
                self.logo_label.setPixmap(scaled)
                self.logo_label.setText("")
            else:
                self.logo_label.setText(tr("Логотип не найден", self.lang))
        else:
            self.logo_label.setPixmap(QPixmap())
            self.logo_label.setText(tr("Логотип не задан", self.lang))

        self.delete_btn.setEnabled(entity in self.user_entities)

    def _export_template(self) -> None:
        entity = self.current_entity()
        if not entity:
            return
        template_path = self.entities.get(entity)
        if not template_path or not os.path.exists(template_path):
            QMessageBox.warning(
                self,
                tr("Ошибка", self.lang),
                tr("Шаблон не найден", self.lang),
            )
            return
        suggested_name = Path(template_path).name
        target, _ = QFileDialog.getSaveFileName(
            self,
            tr("Сохранить шаблон", self.lang),
            suggested_name,
            "Excel (*.xlsx *.xlsm);;All files (*.*)",
        )
        if not target:
            return
        try:
            shutil.copyfile(template_path, target)
        except Exception as exc:
            QMessageBox.critical(
                self,
                tr("Ошибка", self.lang),
                tr("Не удалось сохранить шаблон", self.lang) + f"\n{exc}",
            )

    def _import_template(self) -> None:
        entity = self.current_entity()
        if not entity:
            return
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            tr("Выберите шаблон", self.lang),
            str(Path.home()),
            "Excel (*.xlsx *.xlsm);;All files (*.*)",
        )
        if not file_path:
            return
        self._store_template(entity, file_path)

    def _import_logo(self) -> None:
        entity = self.current_entity()
        if not entity:
            return
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            tr("Выберите логотип", self.lang),
            str(Path.home()),
            "Images (*.png *.jpg *.jpeg *.bmp);;All files (*.*)",
        )
        if not file_path:
            return
        self._store_logo(entity, file_path)

    def _delete_entity(self) -> None:
        entity = self.current_entity()
        if not entity:
            return
        if is_default_entity(entity) and entity not in self.user_entities:
            QMessageBox.information(
                self,
                tr("Информация", self.lang),
                tr("Стандартные юрлица нельзя удалить", self.lang),
            )
            return
        reply = QMessageBox.question(
            self,
            tr("Подтверждение", self.lang),
            tr("Удалить пользовательское юрлицо?", self.lang),
        )
        if reply != QMessageBox.Yes:
            return
        if remove_user_entity(entity):
            self._modified = True
            self.refresh_entities()
        else:
            QMessageBox.critical(
                self,
                tr("Ошибка", self.lang),
                tr("Не удалось удалить запись", self.lang),
            )

    def _add_entity(self) -> None:
        dialog = LegalEntityEditor(self.lang, self)
        if dialog.exec() != QDialog.Accepted:
            return
        data = dialog.get_data()
        if not data:
            return
        name = str(data.get("name", "")).strip()
        if not name:
            return
        if name in self.entities:
            reply = QMessageBox.question(
                self,
                tr("Подтверждение", self.lang),
                tr("Перезаписать существующее юрлицо?", self.lang),
            )
            if reply != QMessageBox.Yes:
                return
        template_path = str(data.get("template"))
        logo_path = data.get("logo")
        vat_enabled = bool(data.get("vat_enabled"))
        default_vat = data.get("default_vat")
        stored_template = self._copy_template_file(name, template_path)
        stored_logo = self._copy_logo_file(name, logo_path) if logo_path else None
        if not stored_template:
            return
        self._persist_entity(
            name,
            stored_template,
            stored_logo,
            vat_enabled=vat_enabled,
            default_vat=default_vat,
        )

    def _store_template(self, entity: str, source_path: str) -> None:
        stored_path = self._copy_template_file(entity, source_path)
        if not stored_path:
            return
        meta = self.metadata.get(entity, {})
        logo_path = meta.get("logo") if isinstance(meta, dict) else None
        vat_enabled = bool(meta.get("vat_enabled")) if isinstance(meta, dict) else False
        default_vat = meta.get("default_vat") if isinstance(meta, dict) else None
        self._persist_entity(
            entity,
            stored_path,
            logo_path,
            vat_enabled=vat_enabled,
            default_vat=default_vat,
        )

    def _store_logo(self, entity: str, source_path: str) -> None:
        stored_logo = self._copy_logo_file(entity, source_path)
        if not stored_logo:
            return
        meta = self.metadata.get(entity, {})
        vat_enabled = bool(meta.get("vat_enabled")) if isinstance(meta, dict) else False
        default_vat = meta.get("default_vat") if isinstance(meta, dict) else None
        template_path = self.entities.get(entity)
        if not template_path:
            QMessageBox.warning(
                self,
                tr("Ошибка", self.lang),
                tr("Сначала задайте шаблон для юрлица", self.lang),
            )
            return
        self._persist_entity(
            entity,
            template_path,
            stored_logo,
            vat_enabled=vat_enabled,
            default_vat=default_vat,
        )

    def _persist_entity(
        self,
        entity: str,
        template_path: str,
        logo_path: Optional[str],
        *,
        vat_enabled: bool,
        default_vat: Optional[float],
    ) -> None:
        data: Dict[str, object] = {"template": template_path}
        if logo_path:
            data["logo"] = logo_path
        data["vat_enabled"] = bool(vat_enabled)
        if vat_enabled and default_vat is not None:
            data["default_vat"] = float(default_vat)
        elif vat_enabled:
            data["default_vat"] = 0.0

        if not set_user_entity(entity, data):
            QMessageBox.critical(
                self,
                tr("Ошибка", self.lang),
                tr("Не удалось сохранить настройки", self.lang),
            )
            return

        self._modified = True
        self.refresh_entities(selected=entity)

    def _copy_template_file(self, entity: str, source_path: str) -> Optional[str]:
        try:
            source = Path(source_path)
            if not source.exists():
                raise FileNotFoundError(source_path)
            target_dir = get_user_templates_dir()
            suffix = source.suffix or ".xlsx"
            filename = f"{_safe_filename(entity)}{suffix}"
            destination = target_dir / filename
            shutil.copyfile(source, destination)
            return str(destination)
        except Exception as exc:
            QMessageBox.critical(
                self,
                tr("Ошибка", self.lang),
                tr("Не удалось скопировать шаблон", self.lang) + f"\n{exc}",
            )
            return None

    def _copy_logo_file(self, entity: str, source_path: str) -> Optional[str]:
        try:
            source = Path(source_path)
            if not source.exists():
                raise FileNotFoundError(source_path)
            target_dir = get_user_logos_dir()
            filename = f"{_safe_filename(entity)}{source.suffix or '.png'}"
            destination = target_dir / filename
            shutil.copyfile(source, destination)
            return str(destination)
        except Exception as exc:
            QMessageBox.critical(
                self,
                tr("Ошибка", self.lang),
                tr("Не удалось сохранить логотип", self.lang) + f"\n{exc}",
            )
            return None
