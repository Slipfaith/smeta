# app.py
from PySide6.QtWidgets import (
    QWidget, QTabWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QToolButton, QApplication
)
from PySide6.QtCore import Qt, QPoint

from tabs.rate_tab import RateTab
from tabs.log_tab import LogTab
from tabs.memoq_tab import MemoqTab
from utils import TITLE_BAR_STYLE, apply_theme


class TitleBar(QWidget):
    """Custom title bar with close and minimize buttons."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAutoFillBackground(True)
        self.setStyleSheet(TITLE_BAR_STYLE)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(5, 0, 5, 0)

        self.title_label = QLabel(parent.windowTitle() if parent else "")
        layout.addWidget(self.title_label)
        layout.addStretch()

        self.min_button = QToolButton()
        self.min_button.setText("-")
        self.min_button.clicked.connect(parent.showMinimized)
        layout.addWidget(self.min_button)

        self.close_button = QToolButton()
        self.close_button.setText("x")
        self.close_button.clicked.connect(parent.close)
        layout.addWidget(self.close_button)

        self._drag_pos = QPoint()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._drag_pos = event.globalPosition().toPoint() - self.parent().frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton:
            self.parent().move(event.globalPosition().toPoint() - self._drag_pos)
            event.accept()

    def mouseReleaseEvent(self, event):
        self._drag_pos = QPoint()

class RateApp(QWidget):
    """
    Главное окно приложения, содержащее вкладки:
      - Ставки (RateTab)
      - Лог (LogTab)
      - MemoQ (MemoqTab)
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Программа для управления ставками")
        self.setGeometry(100, 100, 1200, 800)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Window)

        app = QApplication.instance()
        if app is not None:
            apply_theme(app)

        self.main_layout = QVBoxLayout()
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.setLayout(self.main_layout)

        self.title_bar = TitleBar(self)
        self.main_layout.addWidget(self.title_bar)

        self.tab_widget = QTabWidget()
        self.main_layout.addWidget(self.tab_widget)

        # 1) Вкладка "Ставки"
        self.rate_tab = RateTab()
        self.tab_widget.addTab(self.rate_tab, "Ставки")

        # 2) Вкладка "Лог"
        self.log_tab = LogTab()
        self.tab_widget.addTab(self.log_tab, "Лог")

        # 3) Вкладка "MemoQ"
        self.memoq_tab = MemoqTab()
        self.tab_widget.addTab(self.memoq_tab, "MemoQ")

