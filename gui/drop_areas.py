import os
import traceback
from typing import Callable, Iterable, List, Sequence

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QGroupBox, QMessageBox, QScrollArea


class DropArea(QScrollArea):
    """A scroll area that accepts dropped XML files and forwards them to a callback."""

    def __init__(self, callback: Callable[[Sequence[str]], None], parent=None):
        super().__init__(parent)
        self._callback = callback
        self.setAcceptDrops(True)
        self.setWidgetResizable(True)

        self._base_style = """
            QScrollArea {
                border: 2px dashed #e5e7eb;
                border-radius: 8px;
                background-color: #fafafa;
            }
            QScrollArea[dragOver="true"] {
                border: 2px dashed #2563eb;
                background-color: #eff6ff;
            }
        """
        self.setStyleSheet(self._base_style)

    def disable_hint_style(self):
        self.setStyleSheet(
            """
            QScrollArea[dragOver="true"] {
                border: 2px dashed #2563eb;
                background-color: #eff6ff;
            }
        """
        )

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            all_paths = []
            xml_paths = []
            for url in urls:
                path = url.toLocalFile()
                all_paths.append(path)
                if path.lower().endswith(".xml") or path.lower().endswith(".XML"):
                    xml_paths.append(path)
            if xml_paths:
                event.acceptProposedAction()
                self.setProperty("dragOver", True)
                self.style().unpolish(self)
                self.style().polish(self)
                return
        event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.setProperty("dragOver", False)
        self.style().unpolish(self)
        self.style().polish(self)

    def dropEvent(self, event):
        self.setProperty("dragOver", False)
        self.style().unpolish(self)
        self.style().polish(self)

        if not event.mimeData().hasUrls():
            event.ignore()
            return

        urls = event.mimeData().urls()

        all_paths = []
        xml_paths = []

        for url in urls:
            path = url.toLocalFile()
            all_paths.append(path)

            try:
                if not os.path.exists(path) or not os.path.isfile(path):
                    continue
            except Exception:
                continue

            if path.lower().endswith((".xml", ".XML")):
                xml_paths.append(path)
            else:
                try:
                    with open(path, "r", encoding="utf-8", errors="ignore") as f:
                        first_line = f.readline().strip()
                        if first_line.startswith("<?xml") or "<" in first_line:
                            xml_paths.append(path)
                except Exception:
                    pass

        if xml_paths:
            try:
                self._callback(xml_paths)
                event.acceptProposedAction()
            except Exception as e:
                QMessageBox.critical(
                    None, "Ошибка", f"Ошибка при обработке файлов: {e}"
                )
        else:
            if all_paths:
                QMessageBox.warning(
                    None,
                    "Предупреждение",
                    f"Среди {len(all_paths)} перетащенных файлов не найдено ни одного XML файла.\n",
                    "Поддерживаются только файлы с расширением .xml",
                )
            event.ignore()


class ProjectInfoDropArea(QGroupBox):
    """A group box that accepts dropped Outlook .msg files with project metadata."""

    def __init__(self, title: str, callback: Callable[[Iterable[str]], None], parent=None):
        super().__init__(title, parent)
        self._callback = callback
        self.setAcceptDrops(True)

    def _set_drag_state(self, active: bool):
        self.setProperty("dragOver", active)
        self.style().unpolish(self)
        self.style().polish(self)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                path = url.toLocalFile()
                if path.lower().endswith(".msg"):
                    event.acceptProposedAction()
                    self._set_drag_state(True)
                    return
        event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self._set_drag_state(False)

    def dropEvent(self, event):
        self._set_drag_state(False)
        if not event.mimeData().hasUrls():
            event.ignore()
            return

        msg_paths: List[str] = []
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(".msg"):
                msg_paths.append(path)

        if not msg_paths:
            event.ignore()
            return

        try:
            self._callback(msg_paths)
            event.acceptProposedAction()
        except Exception as exc:
            traceback.print_exc()
            QMessageBox.critical(
                self,
                "Ошибка обработки Outlook файла",
                str(exc) or "Не удалось обработать перетащенный .msg файл.",
            )
            event.ignore()
