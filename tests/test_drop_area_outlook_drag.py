import os
from types import SimpleNamespace

import pytest

try:
    from gui import drop_areas
except ImportError as exc:  # pragma: no cover - platform dependent
    pytest.skip(f"PySide6 is not available: {exc}", allow_module_level=True)

from PySide6.QtCore import QByteArray, QMimeData, QUrl


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


def test_mime_detection_normalises_variants():
    mime = QMimeData()

    descriptor_format = ' application/x-qt-windows-mime;value="FileGroupDescriptor"\x00 '
    contents_format = 'application/x-qt-windows-mime;value="FileContents";Index=0'

    descriptor_bytes = _make_descriptor_bytes("Email.msg", wide=False)
    mime.setData(descriptor_format, QByteArray(descriptor_bytes))
    mime.setData(contents_format, QByteArray(b"message-bytes"))

    assert drop_areas._mime_has_outlook_messages(mime) is True


def test_extract_outlook_messages_creates_temp_files():
    mime = QMimeData()

    descriptor_format = 'application/x-qt-windows-mime;value="FileGroupDescriptorW"'
    contents_format = 'application/x-qt-windows-mime;value="FileContents";Index=0'

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


def test_extract_outlook_messages_reads_via_retrieve_data():
    descriptor_format = 'application/x-qt-windows-mime;value="FileGroupDescriptorW"'
    contents_format = 'application/x-qt-windows-mime;value="FileContents";index=0'

    file_bytes = b"RetrieveData Outlook contents"
    descriptor_bytes = _make_descriptor_bytes("Client Email.msg")

    class RetrieveOnlyMime:
        def formats(self):
            return [descriptor_format, contents_format]

        def data(self, fmt):
            if fmt == descriptor_format:
                return QByteArray(descriptor_bytes)
            if fmt == contents_format:
                return QByteArray()
            return QByteArray()

        def retrieveData(self, fmt, target_type):
            if fmt == descriptor_format:
                return QByteArray(descriptor_bytes)
            if fmt == contents_format:
                if target_type in (bytes, bytearray):
                    return bytes(file_bytes)
                return QByteArray(file_bytes)
            return QByteArray()

        def hasUrls(self):
            return False

        def urls(self):
            return []

    mime = RetrieveOnlyMime()

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


def test_extract_outlook_messages_prefers_existing_msg_files(tmp_path):
    msg_path = tmp_path / "message.msg"
    msg_path.write_bytes(b"existing-msg")

    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(msg_path))])

    created_paths = drop_areas._extract_outlook_messages(mime)

    expected = os.path.normpath(str(msg_path))
    assert [os.path.normpath(path) for path in created_paths] == [expected]


def test_extract_outlook_messages_com_fallback(monkeypatch, tmp_path):
    descriptor_format = 'application/x-qt-windows-mime;value="FileGroupDescriptorW"'
    contents_format = 'application/x-qt-windows-mime;value="FileContents"'

    descriptor_bytes = _make_descriptor_bytes("Client Email.msg")

    mime = QMimeData()
    mime.setData(descriptor_format, QByteArray(descriptor_bytes))
    mime.setData(contents_format, QByteArray())

    monkeypatch.setattr(drop_areas, "_try_qt_filecontents_bytes", lambda *a, **k: b"")
    monkeypatch.setattr(drop_areas, "_parse_ren_private_messages", lambda mime: [(b"entry-id", None)])

    def fake_save_msg(eid, store):
        path = tmp_path / "from_com.msg"
        path.write_bytes(b"com-bytes")
        drop_areas._OUTLOOK_TEMP_FILES.add(str(path))
        return str(path)

    monkeypatch.setattr(drop_areas, "_save_msg_via_outlook_com", fake_save_msg)

    created_paths = drop_areas._extract_outlook_messages(mime)

    try:
        assert created_paths == [str(tmp_path / "from_com.msg")]
        assert os.path.exists(created_paths[0])
        with open(created_paths[0], "rb") as fh:
            assert fh.read() == b"com-bytes"
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
        assert len(created_paths) == 1
        saved_path = created_paths[0]
        assert os.path.exists(saved_path)
        with open(saved_path, "rb") as fh:
            assert fh.read() == contents[0]
    finally:
        for path in created_paths:
            if os.path.exists(path):
                os.remove(path)
            drop_areas._OUTLOOK_TEMP_FILES.discard(path)


def test_collect_existing_msg_paths_ignores_missing_files(tmp_path):
    mime = QMimeData()
    missing_path = tmp_path / "missing.msg"
    mime.setUrls([QUrl.fromLocalFile(str(missing_path))])

    assert drop_areas._collect_existing_msg_paths(mime) == []


def test_collect_existing_msg_paths_returns_existing_files(tmp_path):
    existing_file = tmp_path / "mail.msg"
    existing_file.write_bytes(b"msg data")

    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(existing_file))])

    expected = os.path.normpath(str(existing_file))
    result = [os.path.normpath(path) for path in drop_areas._collect_existing_msg_paths(mime)]
    assert result == [expected]
