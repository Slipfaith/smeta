"""Stylized window for managing memes."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

from PySide6.QtCore import QEasingCurve, QEvent, QPoint, QPropertyAnimation, QSize, Qt
from PySide6.QtGui import QColor, QFont, QIcon, QLinearGradient, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QFrame,
    QGraphicsDropShadowEffect,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from logic.translation_config import tr


@dataclass
class MemeItem:
    """Simple container for meme metadata."""

    title: str
    description: str = ""


def _create_clown_icon(size: int = 64) -> QIcon:
    """Create an icon with a clown emoji rendered in the center."""

    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing)
    font = QFont()
    font.setPointSizeF(size * 0.55)
    painter.setFont(font)
    painter.setPen(Qt.NoPen)
    painter.setBrush(Qt.NoBrush)
    painter.drawText(pixmap.rect(), Qt.AlignCenter, "🤡")
    painter.end()
    return QIcon(pixmap)


class GlassFrame(QFrame):
    """Frame with custom painting to emulate translucent glass."""

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._hovered = False
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAutoFillBackground(False)

    def set_hover(self, value: bool) -> None:
        if self._hovered != value:
            self._hovered = value
            self.update()

    def enterEvent(self, event: QEvent) -> None:  # noqa: D401 - Qt override
        self.set_hover(True)
        super().enterEvent(event)

    def leaveEvent(self, event: QEvent) -> None:  # noqa: D401 - Qt override
        self.set_hover(False)
        super().leaveEvent(event)

    def paintEvent(self, event) -> None:  # noqa: D401 - Qt override
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        rect = self.rect().adjusted(1, 1, -1, -1)
        gradient = QLinearGradient(rect.topLeft(), rect.bottomRight())
        if self._hovered:
            gradient.setColorAt(0.0, QColor(255, 255, 255, 215))
            gradient.setColorAt(0.6, QColor(210, 235, 255, 180))
            gradient.setColorAt(1.0, QColor(180, 220, 250, 160))
        else:
            gradient.setColorAt(0.0, QColor(255, 255, 255, 195))
            gradient.setColorAt(0.6, QColor(215, 235, 255, 165))
            gradient.setColorAt(1.0, QColor(185, 215, 245, 145))

        painter.setPen(QPen(QColor(255, 255, 255, 180), 1.5))
        painter.setBrush(gradient)
        painter.drawRoundedRect(rect, 28, 28)

        highlight = QLinearGradient(rect.topLeft(), rect.topRight())
        highlight.setColorAt(0.0, QColor(255, 255, 255, 110))
        highlight.setColorAt(0.5, QColor(255, 255, 255, 40))
        highlight.setColorAt(1.0, QColor(255, 255, 255, 110))

        painter.setPen(QPen(QColor(255, 255, 255, 70), 2))
        painter.setBrush(Qt.NoBrush)
        painter.drawRoundedRect(rect.adjusted(3, 3, -3, -3), 23, 23)
        painter.end()


class MemeEditorDialog(QDialog):
    """Dialog for editing meme metadata."""

    def __init__(self, item: MemeItem, lang: str, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._item = item
        self.lang = lang
        self.setModal(True)
        self.setWindowTitle(tr("Редактор мемов", lang))
        self.setWindowIcon(_create_clown_icon())
        self.setMinimumSize(360, 280)

        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        self.name_label = QLabel(tr("Название мема", lang))
        layout.addWidget(self.name_label)

        self.name_edit = QLineEdit(item.title)
        layout.addWidget(self.name_edit)

        self.description_label = QLabel(tr("Описание", lang))
        layout.addWidget(self.description_label)

        self.description_edit = QTextEdit(item.description)
        layout.addWidget(self.description_edit, 1)

        self.buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        self.buttons.accepted.connect(self._accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)

        self.set_language(lang)

    def _accept(self) -> None:
        self._item.title = self.name_edit.text().strip() or self._item.title
        self._item.description = self.description_edit.toPlainText().strip()
        self.accept()

    def set_language(self, lang: str) -> None:
        self.lang = lang
        self.setWindowTitle(tr("Редактор мемов", lang))
        self.name_label.setText(tr("Название мема", lang))
        self.description_label.setText(tr("Описание", lang))
        self.buttons.button(QDialogButtonBox.Save).setText(tr("Сохранить", lang))
        self.buttons.button(QDialogButtonBox.Cancel).setText(tr("Отмена", lang))


class MemeWindow(QDialog):
    """Main window for browsing and editing memes."""

    def __init__(self, lang: str = "ru", parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.lang = lang
        self._editor_dialog: Optional[MemeEditorDialog] = None

        self.setWindowTitle(tr("Мемы", lang))
        self.setWindowIcon(_create_clown_icon())
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setMinimumSize(440, 520)

        self._setup_ui()
        self._setup_effects()
        self._populate_initial_items()

    def _setup_ui(self) -> None:
        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)

        self.glass_frame = GlassFrame()
        self.glass_frame.setObjectName("glassFrame")
        outer_layout.addWidget(self.glass_frame)

        frame_layout = QVBoxLayout(self.glass_frame)
        frame_layout.setContentsMargins(32, 32, 32, 32)
        frame_layout.setSpacing(20)

        header_layout = QHBoxLayout()
        header_layout.setSpacing(12)

        self.title_label = QLabel(tr("Мемы", self.lang))
        title_font = self.title_label.font()
        title_font.setPointSize(title_font.pointSize() + 4)
        title_font.setBold(True)
        self.title_label.setFont(title_font)
        header_layout.addWidget(self.title_label)
        header_layout.addStretch(1)

        self.editor_button = QPushButton(tr("Редактировать", self.lang))
        clown_icon = _create_clown_icon(48)
        self.editor_button.setIcon(clown_icon)
        self.editor_button.setIconSize(QSize(48, 48))
        self.editor_button.setCursor(Qt.PointingHandCursor)
        self.editor_button.clicked.connect(self._edit_selected_meme)
        self.editor_button.setToolTip(tr("Редактировать мем", self.lang))
        header_layout.addWidget(self.editor_button)

        frame_layout.addLayout(header_layout)

        self.meme_list = QListWidget()
        self.meme_list.setAlternatingRowColors(True)
        self.meme_list.setStyleSheet(
            """
            QListWidget {
                background: rgba(255, 255, 255, 110);
                border: 1px solid rgba(255, 255, 255, 90);
                border-radius: 18px;
                padding: 12px;
            }
            QListWidget::item {
                padding: 10px 14px;
                border-radius: 12px;
                color: #0f2e52;
            }
            QListWidget::item:selected {
                background: rgba(15, 62, 144, 180);
                color: white;
            }
            """
        )
        self.meme_list.itemDoubleClicked.connect(self._edit_selected_meme)
        self.meme_list.itemSelectionChanged.connect(self._update_actions)
        frame_layout.addWidget(self.meme_list, 1)

        controls_layout = QHBoxLayout()
        controls_layout.setSpacing(12)

        button_style = (
            """
            QPushButton {
                font-size: 20px;
                font-weight: 600;
                min-width: 64px;
                min-height: 48px;
                color: #0f2e52;
                background: rgba(255, 255, 255, 150);
                border: 1px solid rgba(255, 255, 255, 170);
                border-radius: 20px;
            }
            QPushButton:hover {
                background: rgba(245, 250, 255, 210);
            }
            QPushButton:pressed {
                background: rgba(210, 230, 255, 210);
            }
            """
        )

        self.add_button = QPushButton("+")
        self.add_button.setToolTip(tr("Добавить мем", self.lang))
        self.add_button.setCursor(Qt.PointingHandCursor)
        self.add_button.clicked.connect(self._add_meme)
        self.add_button.setStyleSheet(button_style)
        controls_layout.addWidget(self.add_button)

        self.delete_button = QPushButton("🗑️")
        self.delete_button.setToolTip(tr("Удалить выбранные мемы", self.lang))
        self.delete_button.setCursor(Qt.PointingHandCursor)
        self.delete_button.clicked.connect(self._delete_selected)
        self.delete_button.setStyleSheet(button_style)
        controls_layout.addWidget(self.delete_button)

        controls_layout.addStretch(1)
        frame_layout.addLayout(controls_layout)

        self._update_actions()

    def _setup_effects(self) -> None:
        self.shadow_effect = QGraphicsDropShadowEffect(self.glass_frame)
        self.shadow_effect.setBlurRadius(24)
        self.shadow_effect.setOffset(QPoint(0, 12))
        self.shadow_effect.setColor(QColor(15, 52, 96, 90))
        self.glass_frame.setGraphicsEffect(self.shadow_effect)

        self._shadow_animation = QPropertyAnimation(self.shadow_effect, b"blurRadius", self)
        self._shadow_animation.setDuration(250)
        self._shadow_animation.setEasingCurve(QEasingCurve.InOutCubic)

    def _populate_initial_items(self) -> None:
        starters = [
            MemeItem(title=tr("Новый мем", self.lang)),
            MemeItem(title=tr("Смешной мем", self.lang), description=tr("Добавьте описание", self.lang)),
        ]
        for item in starters:
            self._append_item(item)
        if self.meme_list.count():
            self.meme_list.setCurrentRow(0)
        self._update_actions()

    def _append_item(self, item: MemeItem) -> None:
        list_item = QListWidgetItem(item.title)
        list_item.setData(Qt.UserRole, item)
        self.meme_list.addItem(list_item)

    def _current_item(self) -> Optional[QListWidgetItem]:
        return self.meme_list.currentItem()

    def _add_meme(self) -> None:
        item = MemeItem(title=tr("Новый мем", self.lang))
        self._append_item(item)
        last = self.meme_list.item(self.meme_list.count() - 1)
        self.meme_list.setCurrentItem(last)
        self._edit_selected_meme()

    def _delete_selected(self) -> None:
        for list_item in self.meme_list.selectedItems():
            row = self.meme_list.row(list_item)
            self.meme_list.takeItem(row)
        self._update_actions()

    def _edit_selected_meme(self) -> None:
        list_item = self._current_item()
        if list_item is None:
            return
        meme_item: MemeItem = list_item.data(Qt.UserRole)
        if meme_item is None:
            meme_item = MemeItem(title=list_item.text())
        dialog = MemeEditorDialog(meme_item, self.lang, self)
        self._editor_dialog = dialog
        if dialog.exec() == QDialog.Accepted:
            list_item.setText(meme_item.title)
            list_item.setData(Qt.UserRole, meme_item)
        self._update_actions()

    def set_language(self, lang: str) -> None:
        old_lang = self.lang
        self.lang = lang
        self.setWindowTitle(tr("Мемы", lang))
        self.title_label.setText(tr("Мемы", lang))
        self.editor_button.setText(tr("Редактировать", lang))
        self.editor_button.setToolTip(tr("Редактировать мем", lang))
        self.add_button.setToolTip(tr("Добавить мем", lang))
        self.delete_button.setToolTip(tr("Удалить выбранные мемы", lang))
        for index in range(self.meme_list.count()):
            item = self.meme_list.item(index)
            meme_item: MemeItem = item.data(Qt.UserRole)
            if meme_item is None:
                continue
            if meme_item.title == tr("Новый мем", old_lang):
                meme_item.title = tr("Новый мем", lang)
                item.setText(meme_item.title)
            elif meme_item.title == tr("Смешной мем", old_lang):
                meme_item.title = tr("Смешной мем", lang)
                item.setText(meme_item.title)

        if self._editor_dialog is not None:
            self._editor_dialog.set_language(lang)

    def enterEvent(self, event: QEvent) -> None:  # noqa: D401 - Qt override
        self._animate_shadow(42)
        super().enterEvent(event)

    def leaveEvent(self, event: QEvent) -> None:  # noqa: D401 - Qt override
        self._animate_shadow(24)
        super().leaveEvent(event)

    def _animate_shadow(self, radius: float) -> None:
        self._shadow_animation.stop()
        self._shadow_animation.setStartValue(self.shadow_effect.blurRadius())
        self._shadow_animation.setEndValue(radius)
        self._shadow_animation.start()

    def _update_actions(self) -> None:
        has_selection = bool(self.meme_list.selectedItems())
        self.delete_button.setEnabled(has_selection)
        self.editor_button.setEnabled(has_selection)
