import atexit
import locale
import logging
import os
import sys
import tempfile
import traceback
from typing import Callable, Iterable, List, Sequence, Tuple, Optional

from PySide6.QtCore import QByteArray
from PySide6.QtWidgets import QGroupBox, QMessageBox, QScrollArea

from gui.styles import DROP_AREA_BASE_STYLE, DROP_AREA_DRAG_ONLY_STYLE
from logic.translation_config import tr

logger = logging.getLogger(__name__)

# --------------------------- Generic XML DropArea ---------------------------

class DropArea(QScrollArea):
    """Scroll area accepting dropped XML files and forwarding them to a callback."""

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
            xml_paths = []
            for url in urls:
                path = url.toLocalFile()
                if path and path.lower().endswith(".xml"):
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
        xml_paths: List[str] = []
        for url in urls:
            path = url.toLocalFile()
            if not path:
                continue
            if path.lower().endswith(".xml"):
                xml_paths.append(path)
            else:
                try:
                    with open(path, "r", encoding="utf-8", errors="ignore") as f:
                        first_line = f.readline()
                        if "<?xml" in first_line or "<" in first_line:
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
            event.ignore()

# ------------------------- Outlook D&D (MSG) helpers ------------------------

_OUTLOOK_DESCRIPTOR_FORMATS = (
    'application/x-qt-windows-mime;value="FileGroupDescriptorW"',
    'application/x-qt-windows-mime;value="FileGroupDescriptor"',
)
_OUTLOOK_CONTENTS_PREFIX = 'application/x-qt-windows-mime;value="FileContents"'
_FMT_REN_PRIVATE_MESSAGES = 'application/x-qt-windows-mime;value="RenPrivateMessages"'
_FMT_REN_PRIVATE_ITEM = 'application/x-qt-windows-mime;value="RenPrivateItem"'
_FMT_OBJECT_DESCRIPTOR = 'application/x-qt-windows-mime;value="Object Descriptor"'

_OUTLOOK_TEMP_FILES: set[str] = set()
atexit.register(lambda: [os.path.exists(p) and os.remove(p) for p in list(_OUTLOOK_TEMP_FILES)])

def _normalize_mime_format(fmt: str) -> str:
    fmt = fmt.replace("\x00", "").strip()
    return " ".join(fmt.split())

def _iter_mime_format_strings(mime) -> List[str]:
    formats: List[str] = []
    for fmt in mime.formats():
        try:
            decoded = fmt if isinstance(fmt, str) else bytes(fmt).decode("utf-8", errors="ignore")
        except Exception:
            decoded = str(fmt)
        decoded = _normalize_mime_format(decoded)
        if decoded:
            formats.append(decoded)
    logger.info("Outlook drag MIME formats decoded: %s", formats)
    return formats

def _match_descriptor_format(fmt: str) -> bool:
    if fmt in _OUTLOOK_DESCRIPTOR_FORMATS:
        return True
    return "filegroupdescriptor" in fmt.lower()

def _match_file_contents_format(fmt: str) -> bool:
    return "filecontents" in fmt.lower()

def _mime_has_outlook_messages(mime) -> bool:
    if sys.platform != "win32":
        return False
    formats = set(_iter_mime_format_strings(mime))
    return any(_match_descriptor_format(f) for f in formats) and any(_match_file_contents_format(f) for f in formats or [])

def _ensure_bytes(data) -> bytes:
    if not data:
        return b""
    if isinstance(data, (bytes, bytearray)):
        return bytes(data)
    if isinstance(data, QByteArray):
        return bytes(data)
    try:
        return bytes(data)
    except Exception:
        return b""

def _try_qt_filecontents_bytes(mime, formats: Sequence[str], index: int = 0) -> bytes:
    """Try to pull bytes of CFSTR_FILECONTENTS via Qt. Returns b'' if not possible."""
    # Prefer exact variants first
    candidates = []
    exact0 = f'{_OUTLOOK_CONTENTS_PREFIX};index={index}'
    if exact0 in formats:
        candidates.append(exact0)
    if _OUTLOOK_CONTENTS_PREFIX in formats:
        candidates.append(_OUTLOOK_CONTENTS_PREFIX)
    # all other present variants
    for f in formats:
        fl = f.lower()
        if fl.startswith('application/x-qt-windows-mime;value="filecontents"') and f not in candidates:
            candidates.append(f)
    # synthesize common spellings
    for v in (f'{_OUTLOOK_CONTENTS_PREFIX};Index={index}',
              f'{_OUTLOOK_CONTENTS_PREFIX};INDEX={index}'):
        if v not in candidates:
            candidates.append(v)

    logger.debug("Candidate FileContents formats (ordered): %s", candidates)

    # Try data() and retrieveData() for each candidate
    for fmt in candidates:
        try:
            raw = mime.data(fmt)
            b = _ensure_bytes(raw)
            logger.debug("Read %d bytes via data(%s)", len(b), fmt)
            if b:
                logger.info("Obtained %d bytes from FileContents using %s", len(b), fmt)
                return b
        except Exception:
            pass
        for t in (QByteArray, bytes, bytearray):
            try:
                raw = mime.retrieveData(fmt, t)
                b = _ensure_bytes(raw)
                logger.debug("Read %d bytes via retrieveData(%s, %s)", len(b), fmt, getattr(t, "__name__", str(t)))
                if b:
                    logger.info("Obtained %d bytes from FileContents using %s / %s", len(b), fmt, getattr(t, "__name__", str(t)))
                    return b
            except Exception:
                continue
    logger.warning("Failed to obtain FileContents payload after %d attempts", len(candidates))
    return b""

def _decode_filename_from_descriptor(descriptor_bytes: bytes, wide: bool) -> str:
    # FILEGROUPDESCRIPTOR: first 4 bytes = count, then descriptors
    # внутри entry: имя начинается примерно с offset 72 (ANSI) / 76+ (W)
    enc = "utf-16le" if wide else (locale.getpreferredencoding(False) or "utf-8")
    tail = descriptor_bytes[76:] if wide else descriptor_bytes[72:]
    name = tail.decode(enc, errors="ignore").split("\x00", 1)[0].strip()
    return name or "outlook_message.msg"

def _read_descriptor(mime, descriptor_format: str) -> Tuple[int, int, bytes, bool]:
    """Return (count, entry_size, descriptor_bytes, is_wide)."""
    raw = _ensure_bytes(mime.data(descriptor_format)) or _ensure_bytes(mime.retrieveData(descriptor_format, QByteArray))
    logger.debug("Read %d bytes via data(%s)", len(raw), descriptor_format)
    if len(raw) < 4:
        return 0, 0, raw, descriptor_format.endswith("DescriptorW")
    is_wide = descriptor_format.endswith("DescriptorW")
    count = int.from_bytes(raw[:4], "little")
    # пробуем оценить размер записи
    remaining = len(raw) - 4
    entry_size = remaining // count if count else 0
    if entry_size < 72:
        entry_size = 592 if is_wide else 332
    return count, entry_size, raw, is_wide

def _collect_existing_msg_paths(mime) -> List[str]:
    if not mime.hasUrls():
        return []
    paths: List[str] = []
    for url in mime.urls():
        path = url.toLocalFile()
        if path and path.lower().endswith(".msg") and os.path.isfile(path):
            paths.append(path)
    return paths

# --------------------------- COM fallback (MAPI) ---------------------------

def _read_bytes(mime, fmt: str) -> bytes:
    return _ensure_bytes(mime.data(fmt)) or _ensure_bytes(mime.retrieveData(fmt, QByteArray))

def _parse_ren_private_messages(mime) -> List[Tuple[bytes, Optional[bytes]]]:
    """
    Parse RenPrivateMessages / RenPrivateItem to extract (entry_id, store_id).
    Форматы внутренние, но на практике часто идут как:
      [count:DWORD] [ [cbEID:DWORD][EID bytes] [cbStore:DWORD][Store bytes] ] * count
    Возвращает список пар (EntryID, StoreID or None).
    """
    for fmt in (_FMT_REN_PRIVATE_MESSAGES, _FMT_REN_PRIVATE_ITEM):
        data = _read_bytes(mime, fmt)
        if not data:
            continue
        try:
            pos = 0
            if len(data) < 4:
                continue
            count = int.from_bytes(data[pos:pos+4], "little"); pos += 4
            result: List[Tuple[bytes, Optional[bytes]]] = []
            for _ in range(max(1, count)):
                if pos + 4 > len(data): break
                cb_eid = int.from_bytes(data[pos:pos+4], "little"); pos += 4
                if cb_eid <= 0 or pos + cb_eid > len(data): break
                eid = data[pos:pos+cb_eid]; pos += cb_eid
                store: Optional[bytes] = None
                if pos + 4 <= len(data):
                    cb_store = int.from_bytes(data[pos:pos+4], "little"); pos += 4
                    if cb_store > 0 and pos + cb_store <= len(data):
                        store = data[pos:pos+cb_store]; pos += cb_store
                result.append((eid, store))
            if result:
                logger.info("Parsed %d EntryID(s) from %s", len(result), fmt)
                return result
        except Exception as ex:
            logger.warning("Failed to parse %s: %s", fmt, ex)
    return []

def _save_msg_via_outlook_com(eid: bytes, store: Optional[bytes]) -> Optional[str]:
    """
    Use Outlook COM to fetch MailItem by EntryID/StoreID and SaveAs temp .msg
    """
    try:
        import win32com.client  # pywin32
        import pythoncom
    except Exception as ex:
        logger.error("pywin32 not installed; COM fallback unavailable: %s", ex)
        return None

    try:
        pythoncom.CoInitialize()
        app = win32com.client.gencache.EnsureDispatch("Outlook.Application")
        ns = app.GetNamespace("MAPI")
        # EntryID/StoreID должны быть в hex-строке для GetItemFromID
        def b2hex(b: bytes) -> str:
            return "".join(f"{x:02X}" for x in b)
        entry_id = b2hex(eid)
        store_id = b2hex(store) if store else None

        try:
            item = ns.GetItemFromID(entry_id, store_id) if store_id else ns.GetItemFromID(entry_id)
        except Exception as ex:
            logger.error("Namespace.GetItemFromID failed: %s", ex)
            return None

        # 3 == olMSG (прямой .msg)
        fd, temp_path = tempfile.mkstemp(prefix="smeta_outlook_", suffix=".msg")
        os.close(fd)
        item.SaveAs(temp_path, 3)
        _OUTLOOK_TEMP_FILES.add(temp_path)
        logger.info("Saved Outlook item via COM to %s", temp_path)
        return temp_path
    except Exception as ex:
        logger.exception("COM fallback failed: %s", ex)
        return None
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

# ---------------------- Master extractor for Outlook D&D --------------------

def _extract_outlook_messages(mime) -> List[str]:
    """
    Try FileContents (virtual files). If empty (Qt limitation), fallback:
    - parse RenPrivateMessages / RenPrivateItem for EntryID/StoreID
    - fetch item via Outlook COM and SaveAs .msg
    """
    if sys.platform != "win32":
        return []

    formats = _iter_mime_format_strings(mime)

    # 1) If real files already present (.msg dragged from Explorer) — use them.
    real_msgs = _collect_existing_msg_paths(mime)
    if real_msgs:
        return real_msgs

    # 2) Try FileGroupDescriptor -> FileContents (Qt path)
    descriptor_format = next((f for f in formats if _match_descriptor_format(f)), None)
    if descriptor_format:
        count, entry_size, desc_bytes, is_wide = _read_descriptor(mime, descriptor_format)
        if count > 0 and len(desc_bytes) >= 4 + entry_size:
            # one message (по твоему сценарию)
            filename = _decode_filename_from_descriptor(desc_bytes, is_wide)
            payload = _try_qt_filecontents_bytes(mime, formats, index=0)
            if payload:
                fd, temp_path = tempfile.mkstemp(prefix="smeta_outlook_", suffix=os.path.splitext(filename)[1] or ".msg")
                os.close(fd)
                with open(temp_path, "wb") as f:
                    f.write(payload)
                _OUTLOOK_TEMP_FILES.add(temp_path)
                logger.info("Created temporary Outlook file %s (FileContents path)", temp_path)
                return [temp_path]
            else:
                logger.warning("FileContents empty; switching to COM fallback")

    # 3) COM fallback using RenPrivateMessages / RenPrivateItem
    pairs = _parse_ren_private_messages(mime)
    results: List[str] = []
    for (eid, store) in pairs or []:
        path = _save_msg_via_outlook_com(eid, store)
        if path:
            results.append(path)
    if results:
        return results

    # 4) As a very last resort, try ActiveExplorer.Selection[1] (heuristic)
    try:
        import win32com.client  # pywin32
        import pythoncom
        pythoncom.CoInitialize()
        app = win32com.client.gencache.EnsureDispatch("Outlook.Application")
        expl = app.ActiveExplorer()
        if expl and expl.Selection and expl.Selection.Count >= 1:
            item = expl.Selection.Item(1)
            fd, temp_path = tempfile.mkstemp(prefix="smeta_outlook_", suffix=".msg")
            os.close(fd)
            item.SaveAs(temp_path, 3)
            _OUTLOOK_TEMP_FILES.add(temp_path)
            logger.info("Saved ActiveExplorer selection to %s (heuristic fallback)", temp_path)
            return [temp_path]
    except Exception as ex:
        logger.debug("ActiveExplorer heuristic failed: %s", ex)
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    logger.warning("Unable to extract Outlook message bytes from drag data")
    return []

# --------------------------- ProjectInfoDropArea ----------------------------

class ProjectInfoDropArea(QGroupBox):
    """Accepts dropped Outlook messages; yields temp .msg paths to callback."""

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
        if _collect_existing_msg_paths(mime) or _mime_has_outlook_messages(mime):
            event.acceptProposedAction()
            self._set_drag_state(True)
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        mime = event.mimeData()
        if _collect_existing_msg_paths(mime) or _mime_has_outlook_messages(mime):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self._set_drag_state(False)

    def dropEvent(self, event):
        self._set_drag_state(False)
        mime = event.mimeData()

        # 1) paths from Explorer (real .msg)
        msg_paths = _collect_existing_msg_paths(mime)

        # 2) Outlook virtual files / COM fallback
        if not msg_paths:
            msg_paths = _extract_outlook_messages(mime)

        if not msg_paths:
            event.ignore()
            return

        logger.info("Outlook drop resulted in message paths: %s", msg_paths)
        try:
            self._callback(msg_paths)
            event.acceptProposedAction()
        except Exception as exc:
            lang = self._get_lang()
            QMessageBox.critical(
                self,
                tr("Ошибка обработки Outlook файла", lang),
                str(exc) or tr("Не удалось обработать перетащенный .msg файл.", lang),
            )
            event.ignore()