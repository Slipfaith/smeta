import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest

try:
    from PySide6.QtWidgets import QApplication, QLabel, QPushButton
except ImportError:  # pragma: no cover - optional dependency missing in CI
    pytest.skip("PySide6 with Qt widgets is required for UI tests", allow_module_level=True)

from gui.language_pair import LanguagePairWidget
from gui.additional_services import AdditionalServiceTable
from gui.project_setup_widget import ProjectSetupWidget


@pytest.fixture(scope="module")
def qt_app():
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    yield app


def test_language_pair_widget_has_modifiers_button(qt_app):
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


def test_additional_services_table_has_modifiers_button(qt_app):
    table = AdditionalServiceTable()
    try:
        button = getattr(table, "modifiers_button", None)
        assert isinstance(button, QPushButton)
        assert button.text() == "⚙️"
        assert not hasattr(table, "discount_spin")
        assert isinstance(getattr(table, "subtotal_label", None), QLabel)
    finally:
        table.deleteLater()


def test_project_setup_widget_has_modifiers_button(qt_app):
    widget = ProjectSetupWidget()
    try:
        button = getattr(widget, "modifiers_button", None)
        assert isinstance(button, QPushButton)
        assert button.text() == "⚙️"
        assert not hasattr(widget, "discount_spin")
        assert isinstance(getattr(widget, "subtotal_label", None), QLabel)
    finally:
        widget.deleteLater()
