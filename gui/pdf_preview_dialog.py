from typing import Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QSlider,
    QDialogButtonBox,
    QWidget,
)
from PySide6.QtPdfWidgets import QPdfView
from PySide6.QtPdf import QPdfDocument


class PdfPreviewDialog(QDialog):
    """Диалог предварительного просмотра PDF с управлением масштабом."""

    def __init__(self, pdf_path: str, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("Предпросмотр PDF")
        self._doc = QPdfDocument(self)
        self._doc.load(pdf_path)
        self._view = QPdfView(self)
        self._view.setDocument(self._doc)
        self._view.setZoomMode(QPdfView.ZoomMode.Custom)
        self._view.setZoomFactor(1.0)

        self._slider = QSlider(Qt.Horizontal, self)
        self._slider.setRange(25, 400)
        self._slider.setValue(100)
        self._slider.valueChanged.connect(self._on_zoom_changed)

        controls = QHBoxLayout()
        controls.addWidget(QLabel("Масштаб:"))
        controls.addWidget(self._slider)

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addWidget(self._view)
        layout.addLayout(controls)
        layout.addWidget(buttons)

    def _on_zoom_changed(self, value: int) -> None:
        self._view.setZoomFactor(value / 100.0)
