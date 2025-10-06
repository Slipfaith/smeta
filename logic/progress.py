from PySide6.QtWidgets import QProgressDialog, QApplication
from PySide6.QtCore import Qt


class Progress:
    """Wrapper around QProgressDialog for export progress."""

    def __init__(self, title: str = "Сохранение...", parent=None) -> None:
        self.dialog = QProgressDialog(parent)
        self.dialog.setWindowTitle(title)
        self.dialog.setWindowModality(Qt.WindowModal)
        self.dialog.setRange(0, 100)
        self.dialog.setMinimumDuration(0)
        self.dialog.setAutoClose(False)
        self.dialog.setAutoReset(False)
        self.dialog.setCancelButton(None)
        self.dialog.setLabelText("Сохранение...")
        self.dialog.resize(self.dialog.sizeHint())
        self.dialog.setValue(0)

    def set_label(self, message: str) -> None:
        self.dialog.setLabelText(message)
        QApplication.processEvents()

    def set_value(self, percent: int) -> None:
        self.dialog.setValue(percent)
        QApplication.processEvents()

    def on_progress(self, percent: int, message: str) -> None:
        self.set_label(message)
        self.set_value(percent)

    def close(self) -> None:
        self.dialog.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()
