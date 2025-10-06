import os
import logging
import tempfile
from typing import Dict, Any

from .excel_exporter import ExcelExporter
from .excel_process import (
    temporary_separators,
    register_excel_instance,
    unregister_excel_instance,
)

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
    except Exception as e:
        logger.exception("PDF export failed")
        print(f"[PdfExporter] Ошибка экспорта PDF: {e}")
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
        register_excel_instance(excel)
        excel.Visible = False
        excel.DisplayAlerts = False

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
            unregister_excel_instance(excel)

    return success
