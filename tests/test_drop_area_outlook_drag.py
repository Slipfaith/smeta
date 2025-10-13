import os
from types import SimpleNamespace

import pytest

try:
    from gui import drop_areas
except ImportError as exc:  # pragma: no cover - platform dependent
    pytest.skip(f"PySide6 is not available: {exc}", allow_module_level=True)

from PySide6.QtCore import QByteArray, QMimeData


def _make_descriptor_bytes(filenames, wide: bool = True) -> bytes:
    if isinstance(filenames, str):
        filenames = [filenames]

    entry_size = 592 if wide else 332
    filename_offset = 72
    filename_size = 520 if wide else 260
    header = bytearray(4 + entry_size * len(filenames))
    header[:4] = (len(filenames)).to_bytes(4, "little")

    encoding = "utf-16le" if wide else "utf-8"

    offset = 4
    for name in filenames:
        entry = bytearray(entry_size)
        name_bytes = name.encode(encoding)
        entry[filename_offset : filename_offset + min(len(name_bytes), filename_size)] = name_bytes[
            :filename_size
        ]
        header[offset : offset + entry_size] = entry
        offset += entry_size

    return bytes(header)


def _cleanup_paths(paths):
    for path in paths:
        if os.path.exists(path):
            os.remove(path)
        drop_areas._OUTLOOK_TEMP_FILES.discard(path)


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


def test_mime_detection_handles_additional_parameters():
    mime = QMimeData()

    descriptor_format = 'application/x-qt-windows-mime;value="FileGroupDescriptorW";ms=1'
    contents_format = 'application/x-qt-windows-mime;value="FileContents";index=0;ms=1'

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
        _cleanup_paths(created_paths)


def test_extract_outlook_messages_handles_parameterised_formats():
    mime = QMimeData()

    descriptor_format = 'application/x-qt-windows-mime;value="FileGroupDescriptorW";ms=1'
    contents_format = 'application/x-qt-windows-mime;value="FileContents";ms=1'

    file_bytes = b"Parameterised Outlook contents"
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
        _cleanup_paths(created_paths)


def test_extract_outlook_messages_handles_multiple_entries():
    mime = QMimeData()

    descriptor_format = 'application/x-qt-windows-mime;value="FileGroupDescriptorW"'
    file_bytes_a = b"First Outlook contents"
    file_bytes_b = b"Second Outlook contents"
    descriptor_bytes = _make_descriptor_bytes(["EmailA.msg", "EmailB.msg"])

    mime.setData(descriptor_format, QByteArray(descriptor_bytes))
    mime.setData(
        'application/x-qt-windows-mime;value="FileContents";index=0',
        QByteArray(file_bytes_a),
    )
    mime.setData(
        'application/x-qt-windows-mime;value="FileContents";index=1',
        QByteArray(file_bytes_b),
    )

    created_paths = drop_areas._extract_outlook_messages(mime)

    try:
        assert len(created_paths) == 2
        saved_a, saved_b = created_paths
        with open(saved_a, "rb") as fh:
            assert fh.read() == file_bytes_a
        with open(saved_b, "rb") as fh:
            assert fh.read() == file_bytes_b
    finally:
        _cleanup_paths(created_paths)


def test_extract_outlook_messages_supports_ansi_descriptor():
    mime = QMimeData()

    descriptor_format = 'application/x-qt-windows-mime;value="FileGroupDescriptor"'
    contents_format = 'application/x-qt-windows-mime;value="FileContents";index=0'

    file_bytes = b"ANSI descriptor contents"
    descriptor_bytes = _make_descriptor_bytes(["Client Email.msg"], wide=False)

    mime.setData(descriptor_format, QByteArray(descriptor_bytes))
    mime.setData(contents_format, QByteArray(file_bytes))

    created_paths = drop_areas._extract_outlook_messages(mime)

    try:
        assert len(created_paths) == 1
        saved_path = created_paths[0]
        with open(saved_path, "rb") as fh:
            assert fh.read() == file_bytes
    finally:
        _cleanup_paths(created_paths)
