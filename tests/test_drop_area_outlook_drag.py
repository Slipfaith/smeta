import os
from types import SimpleNamespace

import pytest

try:
    from gui import drop_areas
except ImportError as exc:  # pragma: no cover - platform dependent
    pytest.skip(f"PySide6 is not available: {exc}", allow_module_level=True)

from PySide6.QtCore import QByteArray, QMimeData


def _make_descriptor_bytes(filename: str) -> bytes:
    entry_size = 592
    header = bytearray(4 + entry_size)
    header[:4] = (1).to_bytes(4, "little")

    name_bytes = filename.encode("utf-16le")
    name_offset = 4 + 72
    header[name_offset : name_offset + len(name_bytes)] = name_bytes
    return bytes(header)


@pytest.fixture(autouse=True)
def force_windows_platform(monkeypatch):
    monkeypatch.setattr(drop_areas, "sys", SimpleNamespace(platform="win32"))


def test_mime_detection_handles_qbytearray_formats():
    mime = QMimeData()

    descriptor_format = 'application/x-qt-windows-mime;value="FileGroupDescriptorW"'
    contents_format = 'application/x-qt-windows-mime;value="FileContents";index=0'

    descriptor_bytes = _make_descriptor_bytes("Email.msg")
    mime.setData(descriptor_format, QByteArray(descriptor_bytes))
    mime.setData(contents_format, QByteArray(b"message-bytes"))

    assert drop_areas._mime_has_outlook_messages(mime) is True


def test_extract_outlook_messages_creates_temp_files():
    mime = QMimeData()

    descriptor_format = 'application/x-qt-windows-mime;value="FileGroupDescriptorW"'
    contents_format = 'application/x-qt-windows-mime;value="FileContents";index=0'

    file_bytes = b"Outlook attachment contents"
    descriptor_bytes = _make_descriptor_bytes("Client Email.msg")

    mime.setData(descriptor_format, QByteArray(descriptor_bytes))
    mime.setData(contents_format, QByteArray(file_bytes))

    created_paths = drop_areas._extract_outlook_messages(mime)

    try:
        assert len(created_paths) == 1
        saved_path = created_paths[0]
        assert os.path.exists(saved_path)
        with open(saved_path, "rb") as fh:
            assert fh.read() == file_bytes
    finally:
        for path in created_paths:
            if os.path.exists(path):
                os.remove(path)
            drop_areas._OUTLOOK_TEMP_FILES.discard(path)
