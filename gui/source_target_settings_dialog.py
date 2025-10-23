from __future__ import annotations

from typing import Dict, Iterable, List, Sequence, Set

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QVBoxLayout,
)

from logic.translation_config import tr


class SourceTargetSettingsDialog(QDialog):
    """Dialog to configure preferred target languages for sources."""

    def __init__(
        self,
        parent,
        languages: Sequence[dict],
        mapping: Dict[str, Iterable[str]],
        gui_lang: str,
        display_ru: bool,
    ) -> None:
        super().__init__(parent)

        self._languages = list(languages)
        self._display_ru = display_ru
        self._gui_lang = gui_lang
        self._current_source: str | None = None
        self._working_mapping: Dict[str, Set[str]] = {
            str(src or ""): {str(t or "") for t in targets}
            for src, targets in mapping.items()
        }

        self.setWindowTitle(tr("Соответствие языков", gui_lang))

        root_layout = QVBoxLayout(self)

        description = QLabel(tr("Выберите целевые языки для каждого исходного.", gui_lang))
        description.setWordWrap(True)
        root_layout.addWidget(description)

        content_layout = QHBoxLayout()
        root_layout.addLayout(content_layout)

        self._source_list = QListWidget()
        self._source_list.currentItemChanged.connect(self._on_source_changed)
        content_layout.addWidget(self._source_list, 1)

        right_layout = QVBoxLayout()
        content_layout.addLayout(right_layout, 2)

        self._search_edit = QLineEdit()
        self._search_edit.setPlaceholderText(tr("Поиск...", gui_lang))
        self._search_edit.textChanged.connect(self._apply_filter)
        right_layout.addWidget(self._search_edit)

        self._target_list = QListWidget()
        self._target_list.setSelectionMode(QListWidget.NoSelection)
        right_layout.addWidget(self._target_list)

        target_controls = QHBoxLayout()
        self._select_all_btn = QPushButton(tr("Выбрать все", gui_lang))
        self._select_all_btn.clicked.connect(lambda: self._set_all_targets(Qt.Checked))
        target_controls.addWidget(self._select_all_btn)
        self._clear_btn = QPushButton(tr("Снять выбор", gui_lang))
        self._clear_btn.clicked.connect(lambda: self._set_all_targets(Qt.Unchecked))
        target_controls.addWidget(self._clear_btn)
        target_controls.addStretch(1)
        right_layout.addLayout(target_controls)

        self._populate_sources()
        self._populate_targets()

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        root_layout.addWidget(buttons)

        if self._source_list.count() > 0:
            self._source_list.setCurrentRow(0)

        self.resize(720, 520)

    def _populate_sources(self) -> None:
        self._source_list.clear()
        for lang in self._languages:
            en_name = str(lang.get("en", ""))
            ru_name = str(lang.get("ru", ""))
            display = ru_name if self._display_ru else en_name
            item = QListWidgetItem(display or en_name or ru_name)
            item.setData(Qt.UserRole, en_name)
            self._source_list.addItem(item)

    def _populate_targets(self) -> None:
        self._target_list.clear()
        for lang in self._languages:
            en_name = str(lang.get("en", ""))
            ru_name = str(lang.get("ru", ""))
            display = ru_name if self._display_ru else en_name
            item = QListWidgetItem(display or en_name or ru_name)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)
            item.setData(Qt.UserRole, {"en": en_name, "ru": ru_name})
            self._target_list.addItem(item)

    def _store_current_selection(self) -> None:
        if not self._current_source:
            return
        selected: Set[str] = set()
        for idx in range(self._target_list.count()):
            item = self._target_list.item(idx)
            if item.checkState() != Qt.Checked:
                continue
            lang = item.data(Qt.UserRole) or {}
            en_name = str(lang.get("en", ""))
            if en_name:
                selected.add(en_name)
        if selected:
            self._working_mapping[self._current_source] = selected
        elif self._current_source in self._working_mapping:
            self._working_mapping.pop(self._current_source)

    def _on_source_changed(self, current: QListWidgetItem, previous: QListWidgetItem | None) -> None:
        if previous is not None:
            self._store_current_selection()

        source_en = current.data(Qt.UserRole) if current is not None else None
        self._current_source = str(source_en or "") or None
        self._apply_mapping_to_targets()

    def _apply_mapping_to_targets(self) -> None:
        selected = self._working_mapping.get(self._current_source or "", set())
        for idx in range(self._target_list.count()):
            item = self._target_list.item(idx)
            lang = item.data(Qt.UserRole) or {}
            en_name = str(lang.get("en", ""))
            item.setCheckState(Qt.Checked if en_name in selected else Qt.Unchecked)

    def _set_all_targets(self, state: Qt.CheckState) -> None:
        for idx in range(self._target_list.count()):
            item = self._target_list.item(idx)
            item.setCheckState(state)

    def _apply_filter(self, text: str) -> None:
        pattern = text.strip().lower()
        for idx in range(self._target_list.count()):
            item = self._target_list.item(idx)
            lang = item.data(Qt.UserRole) or {}
            en_name = str(lang.get("en", "")).lower()
            ru_name = str(lang.get("ru", "")).lower()
            display = item.text().lower()
            match = not pattern or pattern in display or pattern in en_name or pattern in ru_name
            item.setHidden(not match)

    def accept(self) -> None:  # type: ignore[override]
        self._store_current_selection()
        super().accept()

    def result_mapping(self) -> Dict[str, List[str]]:
        """Return mapping with English names preserving selection order."""

        result: Dict[str, List[str]] = {}
        for idx in range(self._source_list.count()):
            item = self._source_list.item(idx)
            src_en = str(item.data(Qt.UserRole) or "")
            selected = self._working_mapping.get(src_en, set())
            if not selected:
                continue
            ordered: List[str] = []
            for t_idx in range(self._target_list.count()):
                lang = self._target_list.item(t_idx).data(Qt.UserRole) or {}
                en_name = str(lang.get("en", ""))
                if en_name in selected and en_name not in ordered:
                    ordered.append(en_name)
            # Add any remaining (e.g. languages not present in list anymore)
            for en_name in selected:
                if en_name not in ordered:
                    ordered.append(en_name)
            if ordered:
                result[src_en] = ordered
        return result
