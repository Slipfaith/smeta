import atexit
import locale
import os
import sys
import tempfile
import traceback
from typing import Callable, Iterable, List, Sequence

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QGroupBox, QMessageBox, QScrollArea

from gui.styles import DROP_AREA_BASE_STYLE, DROP_AREA_DRAG_ONLY_STYLE
from logic.translation_config import tr


class DropArea(QScrollArea):
    """A scroll area that accepts dropped XML files and forwards them to a callback."""

    def __init__(
        self,
        callback: Callable[[Sequence[str]], None],
        get_lang: Callable[[], str] | None = None,
        parent=None,
    ):
        super().__init__(parent)
        self._callback = callback
        self._get_lang = get_lang or (lambda: "ru")
        self.setAcceptDrops(True)
        self.setWidgetResizable(True)

        self._base_style = DROP_AREA_BASE_STYLE
        self.setStyleSheet(self._base_style)

    def disable_hint_style(self):
        self.setStyleSheet(DROP_AREA_DRAG_ONLY_STYLE)

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
                lang = self._get_lang()
                QMessageBox.critical(
                    self.window(),
                    tr("Ошибка", lang),
                    tr("Ошибка при обработке файлов: {0}", lang).format(e),
                )
        else:
            if all_paths:
                lang = self._get_lang()
                QMessageBox.warning(
                    self.window(),
                    tr("Предупреждение", lang),
                    "\n".join(
                        [
                            tr(
                                "Среди {0} перетащенных файлов не найдено ни одного XML файла.",
                                lang,
                            ).format(len(all_paths)),
                            tr(
                                "Поддерживаются только файлы с расширением .xml", lang
                            ),
                        ]
                    ),
                )
            event.ignore()


_OUTLOOK_DESCRIPTOR_FORMATS = (
    'application/x-qt-windows-mime;value="FileGroupDescriptorW"',
    'application/x-qt-windows-mime;value="FileGroupDescriptor"',
)
_OUTLOOK_CONTENTS_PREFIX = 'application/x-qt-windows-mime;value="FileContents"'
_OUTLOOK_TEMP_FILES: set[str] = set()


def _cleanup_temp_outlook_files():
    while _OUTLOOK_TEMP_FILES:
        path = _OUTLOOK_TEMP_FILES.pop()
        try:
            if os.path.exists(path):
                os.remove(path)
        except OSError:
            # Best-effort cleanup. Ignore failures, files will be removed
            # by the OS temp directory cleanup policies.
            pass


atexit.register(_cleanup_temp_outlook_files)


def _normalize_mime_format(fmt: str) -> str:
    """Return *fmt* stripped and with redundant whitespace removed."""

    fmt = fmt.strip()
    # Outlook may append extra parameters separated by semicolons. Duplicated
    # whitespace around those separators changes the literal string returned by
    # Qt, so collapse it to improve comparisons.
    return " ".join(fmt.split())


def _iter_mime_format_strings(mime) -> List[str]:
    """Return all MIME formats from *mime* as decoded strings."""

    formats: List[str] = []
    for fmt in mime.formats():
        if isinstance(fmt, str):
            formats.append(_normalize_mime_format(fmt))
            continue

        if isinstance(fmt, bytes):
            decoded = fmt.decode("utf-8", errors="ignore")
            if decoded:
                formats.append(_normalize_mime_format(decoded))
            continue

        # PySide6 returns QByteArray instances for MIME formats. Converting them
        # to ``bytes`` yields the raw data that we then decode to a string.
        try:
            decoded = bytes(fmt).decode("utf-8", errors="ignore")
        except Exception:
            decoded = str(fmt)

        decoded = _normalize_mime_format(decoded)
        if decoded:
            formats.append(decoded)

    return formats


def _match_descriptor_format(fmt: str) -> bool:
    if fmt in _OUTLOOK_DESCRIPTOR_FORMATS:
        return True

    if not fmt.startswith("application/x-qt-windows-mime;value="):
        return False

    remainder = fmt.split(";value=", 1)[-1]
    if remainder.startswith('"'):
        remainder = remainder[1:]
    value_part = remainder.split('"', 1)[0]
    return value_part in {"FileGroupDescriptor", "FileGroupDescriptorW"}


def _match_file_contents_format(fmt: str) -> bool:
    if fmt == _OUTLOOK_CONTENTS_PREFIX:
        return True

    if not fmt.startswith(_OUTLOOK_CONTENTS_PREFIX):
        return False

    return True


def _mime_has_outlook_messages(mime) -> bool:
    """Return True if *mime* contains Outlook drag-n-drop data."""

    if sys.platform != "win32":
        return False

    formats = set(_iter_mime_format_strings(mime))
    if not any(_match_descriptor_format(fmt) for fmt in formats):
        return False

    # Outlook provides one "FileContents" entry per dragged message.
    if any(_match_file_contents_format(fmt) for fmt in formats):
        return True

    return False


def _find_file_contents_format(formats: Sequence[str], index: int) -> str | None:
    """Return the Qt MIME format that stores data for the requested *index*."""

    exact_name = f"{_OUTLOOK_CONTENTS_PREFIX};index={index}"
    if exact_name in formats:
        return exact_name

    prefix = f"{_OUTLOOK_CONTENTS_PREFIX};"
    for fmt in formats:
        if fmt.startswith(prefix) and "index=" in fmt:
            try:
                index_part = fmt.split("index=", 1)[-1].split(";", 1)[0]
                fmt_index = int(index_part)
            except ValueError:
                continue
            if fmt_index == index:
                return fmt

    for fmt in formats:
        if fmt.startswith(prefix):
            return fmt if index == 0 else None

    if index == 0 and any(fmt == _OUTLOOK_CONTENTS_PREFIX for fmt in formats):
        return _OUTLOOK_CONTENTS_PREFIX

    return None


def _decode_filename(data: bytes, wide: bool) -> str:
    """Decode a filename stored inside a FILEDESCRIPTOR structure."""

    encoding = "utf-16le" if wide else locale.getpreferredencoding(False) or "utf-8"
    decoded = data.decode(encoding, errors="ignore").split("\x00", 1)[0]
    return decoded.strip()


def _extract_outlook_messages(mime) -> List[str]:
    """Return temporary file paths created from the Outlook drag data."""

    if sys.platform != "win32":
        return []

    formats = _iter_mime_format_strings(mime)
    descriptor_format = next(
        (fmt for fmt in formats if _match_descriptor_format(fmt)),
        None,
    )

    if not descriptor_format:
        return []

    descriptor_bytes = bytes(mime.data(descriptor_format))
    if len(descriptor_bytes) < 4:
        return []

    is_wide = descriptor_format.endswith("DescriptorW")
    entry_size = 592 if is_wide else 332
    filename_size = 520 if is_wide else 260

    count = int.from_bytes(descriptor_bytes[:4], "little")
    offset = 4
    created_paths: List[str] = []

    for index in range(count):
        if offset + entry_size > len(descriptor_bytes):
            break

        entry = descriptor_bytes[offset : offset + entry_size]
        offset += entry_size

        filename_bytes = entry[72 : 72 + filename_size]
        filename = _decode_filename(filename_bytes, wide=is_wide)
        if not filename:
            continue

        contents_format = _find_file_contents_format(formats, index)
        if not contents_format:
            continue

        file_bytes = bytes(mime.data(contents_format))
        if not file_bytes:
            continue

        suffix = os.path.splitext(filename)[1] or ".msg"
        prefix = "smeta_outlook_"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix, prefix=prefix) as tmp_file:
            tmp_file.write(file_bytes)
            temp_path = tmp_file.name

        _OUTLOOK_TEMP_FILES.add(temp_path)
        created_paths.append(temp_path)

    return created_paths


class ProjectInfoDropArea(QGroupBox):
    """A group box that accepts dropped Outlook .msg files with project metadata."""

    def __init__(
        self,
        title: str,
        callback: Callable[[Iterable[str]], None],
        get_lang: Callable[[], str] | None = None,
        parent=None,
    ):
        super().__init__(title, parent)
        self._callback = callback
        self._get_lang = get_lang or (lambda: "ru")
        self.setAcceptDrops(True)

    def _set_drag_state(self, active: bool):
        self.setProperty("dragOver", active)
        self.style().unpolish(self)
        self.style().polish(self)

    def dragEnterEvent(self, event):
        mime = event.mimeData()
        if mime.hasUrls():
            for url in mime.urls():
                path = url.toLocalFile()
                if path.lower().endswith(".msg"):
                    event.acceptProposedAction()
                    self._set_drag_state(True)
                    return
        elif _mime_has_outlook_messages(mime):
            event.acceptProposedAction()
            self._set_drag_state(True)
            return
        event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls() or _mime_has_outlook_messages(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self._set_drag_state(False)

    def dropEvent(self, event):
        self._set_drag_state(False)
        msg_paths: List[str] = []
        mime = event.mimeData()

        if mime.hasUrls():
            for url in mime.urls():
                path = url.toLocalFile()
                if path.lower().endswith(".msg"):
                    msg_paths.append(path)

        if not msg_paths:
            msg_paths.extend(_extract_outlook_messages(mime))

        if not msg_paths:
            event.ignore()
            return

        try:
            self._callback(msg_paths)
            event.acceptProposedAction()
        except Exception as exc:
            traceback.print_exc()
            lang = self._get_lang()
            QMessageBox.critical(
                self,
                tr("Ошибка обработки Outlook файла", lang),
                str(exc) or tr("Не удалось обработать перетащенный .msg файл.", lang),
            )
            event.ignore()
