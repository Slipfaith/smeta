import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest

QtWidgets = pytest.importorskip("PySide6.QtWidgets", exc_type=ImportError)
QApplication = QtWidgets.QApplication
QLabel = QtWidgets.QLabel
QPushButton = QtWidgets.QPushButton

from gui.language_pair import LanguagePairWidget


@pytest.fixture(scope="module")
def qt_app():
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    yield app


def test_language_pair_widget_has_modifiers_button(qt_app):
    assert qt_app is not None
    widget = LanguagePairWidget("Test")
    try:
        group = widget.translation_group
        button = getattr(group, "modifiers_button", None)
        assert isinstance(button, QPushButton)
        assert button.text() == "⚙️"
        assert not hasattr(group, "discount_spin")
        assert isinstance(getattr(group, "subtotal_label", None), QLabel)
    finally:
        widget.deleteLater()
