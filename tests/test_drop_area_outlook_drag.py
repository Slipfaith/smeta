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
    header = bytearray(4 + entry_size * len(filenames))
    header[:4] = (len(filenames)).to_bytes(4, "little")

    encoding = "utf-16le" if wide else "utf-8"

    for index, filename in enumerate(filenames):
        name_bytes = filename.encode(encoding)
        name_offset = 4 + index * entry_size + 72
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
        for path in created_paths:
            if os.path.exists(path):
                os.remove(path)
            drop_areas._OUTLOOK_TEMP_FILES.discard(path)


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
        for path in created_paths:
            if os.path.exists(path):
                os.remove(path)
            drop_areas._OUTLOOK_TEMP_FILES.discard(path)


def test_extract_outlook_messages_handles_ansi_descriptor():
    mime = QMimeData()

    descriptor_format = 'application/x-qt-windows-mime;value="FileGroupDescriptor"'
    contents_format = 'application/x-qt-windows-mime;value="FileContents";index=0'

    file_bytes = b"ANSI Outlook contents"
    descriptor_bytes = _make_descriptor_bytes("Client Email.msg", wide=False)

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


def test_extract_outlook_messages_handles_multiple_items():
    mime = QMimeData()

    descriptor_format = 'application/x-qt-windows-mime;value="FileGroupDescriptorW"'
    contents_base = 'application/x-qt-windows-mime;value="FileContents"'

    filenames = ["First.msg", "Second.msg"]
    descriptor_bytes = _make_descriptor_bytes(filenames)

    mime.setData(descriptor_format, QByteArray(descriptor_bytes))

    contents = [b"first-bytes", b"second-bytes"]
    for index, data in enumerate(contents):
        contents_format = f"{contents_base};index={index}"
        mime.setData(contents_format, QByteArray(data))

    created_paths = drop_areas._extract_outlook_messages(mime)

    try:
        assert len(created_paths) == 2
        for path, expected_bytes in zip(created_paths, contents):
            assert os.path.exists(path)
            with open(path, "rb") as fh:
                assert fh.read() == expected_bytes
    finally:
        for path in created_paths:
            if os.path.exists(path):
                os.remove(path)
            drop_areas._OUTLOOK_TEMP_FILES.discard(path)
