from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import QLabel


class ScaledPixmapLabel(QLabel):
    """QLabel that scales pixmaps while preserving aspect ratio."""

    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        self._pixmap: QPixmap | None = None
        self.setAlignment(Qt.AlignCenter)

    def setPixmap(self, pixmap: QPixmap) -> None:  # type: ignore[override]
        self._pixmap = pixmap if pixmap and not pixmap.isNull() else None
        if self._pixmap is None:
            super().setPixmap(QPixmap())
        else:
            super().setPixmap(self._scaled_pixmap())

    def clear(self) -> None:  # type: ignore[override]
        self._pixmap = None
        super().clear()
        super().setPixmap(QPixmap())

    def resizeEvent(self, event) -> None:  # type: ignore[override]
        super().resizeEvent(event)
        if self._pixmap is not None:
            super().setPixmap(self._scaled_pixmap())

    def set_path(self, path: Path | str) -> bool:
        pixmap = QPixmap(str(path))
        if pixmap.isNull():
            self.clear()
            return False
        self.setPixmap(pixmap)
        return True

    def _scaled_pixmap(self) -> QPixmap:
        assert self._pixmap is not None
        return self._pixmap.scaled(self.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
