import os
import logging
import tempfile
from contextlib import contextmanager
from typing import Dict, Any

from .excel_exporter import ExcelExporter
from .excel_process import temporary_separators

logger = logging.getLogger("PdfExporter")


def export_to_pdf(
    project_data: Dict[str, Any],
    output_path: str,
    template_path: str,
    currency: str = "RUB",
    lang: str = "ru",
) -> bool:
    """Export project data to PDF using ExcelExporter and xlsx_to_pdf conversion."""
    try:
        exporter = ExcelExporter(template_path, currency=currency, lang=lang)
        with tempfile.TemporaryDirectory() as tmpdir:
            xlsx_path = os.path.join(tmpdir, "temp.xlsx")
            if not exporter.export_to_excel(project_data, xlsx_path, fit_to_page=True):
                return False
            if not xlsx_to_pdf(xlsx_path, output_path, lang=lang):
                raise RuntimeError("Не удалось конвертировать в PDF")
        return True
    except Exception:
        logger.exception("Ошибка экспорта PDF")
        return False


def xlsx_to_pdf(xlsx_path: str, pdf_path: str, lang: str = "ru") -> bool:
    """Convert an XLSX file to PDF using Excel.

    The Excel automation backend is instructed to use language-specific
    decimal and thousands separators so that numbers in the resulting PDF
    always display with the correct symbols regardless of the system locale.
    """

    excel = wb = None
    try:
        import win32com.client  # type: ignore

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        with _temporary_excel_speedup(excel, win32com.client):
            with temporary_separators(excel, lang):
                wb = excel.Workbooks.Open(xlsx_path)
                wb.ExportAsFixedFormat(0, pdf_path)
                success = os.path.exists(pdf_path)
    except Exception:
        success = False
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

    return success


@contextmanager
def _temporary_excel_speedup(excel, win32com_client):
    """Temporarily disable heavy Excel features to speed up export."""

    original_values: Dict[str, Any] = {}
    to_disable = (
        ("ScreenUpdating", False),
        ("EnableEvents", False),
        ("DisplayStatusBar", False),
    )

    try:
        for attr, target in to_disable:
            try:
                original_values[attr] = getattr(excel, attr)
                setattr(excel, attr, target)
            except Exception:
                pass

        # Switch calculation mode to manual and restore the original value afterwards.
        try:
            manual_constant = getattr(
                getattr(win32com_client, "constants", None),
                "xlCalculationManual",
                -4135,
            )
            original_values["Calculation"] = getattr(excel, "Calculation")
            setattr(excel, "Calculation", manual_constant)
        except Exception:
            original_values.pop("Calculation", None)

        yield
    finally:
        for attr, value in original_values.items():
            try:
                setattr(excel, attr, value)
            except Exception:
                pass
