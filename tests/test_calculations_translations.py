from __future__ import annotations

import importlib
import sys
import types
from contextlib import contextmanager


@contextmanager
def _fake_pyside_modules() -> None:
    names = ["PySide6", "PySide6.QtCore", "PySide6.QtWidgets"]
    saved = {name: sys.modules.pop(name, None) for name in names}

    package = types.ModuleType("PySide6")
    qt_core = types.ModuleType("PySide6.QtCore")
    qt_core.Qt = types.SimpleNamespace(MatchFixedString=0)

    class _FakeQInputDialog:
        @staticmethod
        def getDouble(*_args, **_kwargs):
            return 0.0, False

    qt_widgets = types.ModuleType("PySide6.QtWidgets")
    qt_widgets.QInputDialog = _FakeQInputDialog

    package.QtCore = qt_core
    package.QtWidgets = qt_widgets

    sys.modules["PySide6"] = package
    sys.modules["PySide6.QtCore"] = qt_core
    sys.modules["PySide6.QtWidgets"] = qt_widgets

    try:
        yield
    finally:
        for name in names:
            module = saved[name]
            if module is None:
                sys.modules.pop(name, None)
            else:
                sys.modules[name] = module


class _DummyLabel:
    def __init__(self) -> None:
        self.text_value = ""
        self.visible = False

    def setText(self, text: str) -> None:  # noqa: N802 - Qt style name
        self.text_value = text

    def show(self) -> None:
        self.visible = True

    def hide(self) -> None:
        self.visible = False


class _DummySpinBox:
    def __init__(self, value: float, enabled: bool = False) -> None:
        self._value = value
        self._enabled = enabled

    def value(self) -> float:
        return self._value

    def isEnabled(self) -> bool:  # noqa: N802 - Qt style name
        return self._enabled


class _DummyWidget:
    def __init__(self, subtotal: float, discount: float, markup: float) -> None:
        self._subtotal = subtotal
        self._discount = discount
        self._markup = markup

    def get_subtotal(self) -> float:
        return self._subtotal

    def get_discount_amount(self) -> float:
        return self._discount

    def get_markup_amount(self) -> float:
        return self._markup


class _DummyWindow:
    def __init__(self) -> None:
        self.lang_display_ru = False
        self.currency_symbol = "USD"
        self.vat_spin = _DummySpinBox(0.0, enabled=False)
        self.markup_total_label = _DummyLabel()
        self.discount_total_label = _DummyLabel()
        self.total_label = _DummyLabel()
        self.project_setup_widget = _DummyWidget(100.0, 10.0, 5.0)
        self.language_pairs = {"en-ru": _DummyWidget(200.0, 0.0, 20.0)}
        self.additional_services_widget = _DummyWidget(50.0, 5.0, 0.0)


def test_update_total_uses_english_discount_and_markup_labels():
    sys.modules.pop("logic.calculations", None)
    with _fake_pyside_modules():
        calculations = importlib.import_module("logic.calculations")

    window = _DummyWindow()

    calculations.update_total(window)

    assert window.discount_total_label.visible is True
    assert window.markup_total_label.visible is True
    assert window.discount_total_label.text_value.startswith("Discount amount:")
    assert window.markup_total_label.text_value.startswith("Markup amount:")
    assert window.total_label.text_value.startswith("Total:")
