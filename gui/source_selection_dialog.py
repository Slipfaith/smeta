from __future__ import annotations

from typing import Iterable, List, Sequence

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QVBoxLayout,
)

from logic.translation_config import tr


class SourceSelectionDialog(QDialog):
    """Dialog that allows selecting multiple source languages."""

    def __init__(
        self,
        parent,
        languages: Sequence[dict],
        selected_norms: Iterable[str],
        gui_lang: str,
        display_ru: bool,
    ) -> None:
        super().__init__(parent)
        self._languages = list(languages)
        self._display_ru = display_ru
        self._selected_norms = {str(name or "").strip().lower() for name in selected_norms}
        self._gui_lang = gui_lang

        self.setWindowTitle(tr("Выбор исходных языков", gui_lang))
        layout = QVBoxLayout(self)

        self._search_edit = QLineEdit()
        self._search_edit.setPlaceholderText(tr("Поиск...", gui_lang))
        self._search_edit.textChanged.connect(self._apply_filter)
        layout.addWidget(self._search_edit)

        self._list_widget = QListWidget()
        self._list_widget.setSelectionMode(QListWidget.NoSelection)
        layout.addWidget(self._list_widget)

        controls = QHBoxLayout()
        self._select_all_btn = QPushButton(tr("Выбрать все", gui_lang))
        self._select_all_btn.clicked.connect(lambda: self._set_all(Qt.Checked))
        controls.addWidget(self._select_all_btn)
        self._clear_btn = QPushButton(tr("Снять выбор", gui_lang))
        self._clear_btn.clicked.connect(lambda: self._set_all(Qt.Unchecked))
        controls.addWidget(self._clear_btn)
        controls.addStretch(1)
        layout.addLayout(controls)

        self._populate_list()

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.resize(420, 500)

    def _populate_list(self) -> None:
        self._list_widget.clear()
        for lang in self._languages:
            en_name = str(lang.get("en", ""))
            ru_name = str(lang.get("ru", ""))
            display = ru_name if self._display_ru else en_name
            item = QListWidgetItem(display or en_name or ru_name)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(
                Qt.Checked if en_name.strip().lower() in self._selected_norms else Qt.Unchecked
            )
            item.setData(Qt.UserRole, dict(lang))
            self._list_widget.addItem(item)

    def _set_all(self, state: Qt.CheckState) -> None:
        for idx in range(self._list_widget.count()):
            item = self._list_widget.item(idx)
            item.setCheckState(state)

    def _apply_filter(self, text: str) -> None:
        pattern = text.strip().lower()
        for idx in range(self._list_widget.count()):
            item = self._list_widget.item(idx)
            lang = item.data(Qt.UserRole) or {}
            en_name = str(lang.get("en", "")).lower()
            ru_name = str(lang.get("ru", "")).lower()
            display = item.text().lower()
            match = not pattern or pattern in display or pattern in en_name or pattern in ru_name
            item.setHidden(not match)

    def selected_languages(self) -> List[dict]:
        selected: List[dict] = []
        for idx in range(self._list_widget.count()):
            item = self._list_widget.item(idx)
            if item.checkState() != Qt.Checked:
                continue
            lang = item.data(Qt.UserRole)
            if isinstance(lang, dict):
                selected.append(dict(lang))
        return selected
