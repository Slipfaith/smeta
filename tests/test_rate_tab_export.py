import pytest

try:  # pragma: no cover - environment-specific import guard
    from PySide6.QtWidgets import QApplication, QTableWidgetItem
except ImportError as exc:  # pragma: no cover - executed only when Qt is missing
    pytest.skip(f"PySide6 QtWidgets unavailable: {exc}", allow_module_level=True)

from logic.translation_config import tr
from rates1.tabs.rate_tab import RateTab


@pytest.fixture(scope="module")
def qapp():
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    yield app


def test_rate_tab_dataframe_uses_gui_rows(qapp):
    rate_tab = RateTab(lambda: "en")

    rate_tab.table.setColumnCount(5)
    headers = [
        tr("Исходный язык", "en"),
        tr("Язык перевода", "en"),
        "Basic",
        "Complex",
        "Hour",
    ]
    rate_tab.table.setHorizontalHeaderLabels(headers)
    rate_tab.table.setRowCount(1)
    rate_tab.table.setItem(0, 0, QTableWidgetItem("English"))
    rate_tab.table.setItem(0, 1, QTableWidgetItem("Russian"))
    rate_tab.table.setItem(0, 2, QTableWidgetItem("1.1"))
    rate_tab.table.setItem(0, 3, QTableWidgetItem("2.2"))
    rate_tab.table.setItem(0, 4, QTableWidgetItem("3"))

    rate_tab._selected_targets_by_source = {"English": ["Russian"]}
    rate_tab._source_order = ["English"]

    rate_tab._emit_current_selection()
    rows = rate_tab._export_payloads.get(("USD", 1))
    assert rows is not None

    df = rate_tab._dataframe_from_rows(rows, "en")

    source_col = tr("Исходный язык", "en")
    target_col = tr("Язык перевода", "en")

    assert list(df.columns) == [source_col, target_col, "Basic", "Complex", "Hour"]
    assert df.iloc[0][source_col] == "English"
    assert df.iloc[0][target_col] == "Russian"
    assert df.iloc[0]["Basic"] == pytest.approx(1.1)
    assert df.iloc[0]["Complex"] == pytest.approx(2.2)
    assert df.iloc[0]["Hour"] == pytest.approx(3)

    rate_tab.deleteLater()
