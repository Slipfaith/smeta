# logic/pdf_exporter.py
import os
import logging
import tempfile
import subprocess
from typing import Dict, Any

from .excel_exporter import ExcelExporter

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
            if not xlsx_to_pdf(xlsx_path, output_path):
                raise RuntimeError("Не удалось конвертировать в PDF")
        return True
    except Exception as e:
        logger.exception("PDF export failed")
        print(f"[PdfExporter] Ошибка экспорта PDF: {e}")
        return False


def xlsx_to_pdf(xlsx_path: str, pdf_path: str) -> bool:
    """Convert XLSX file to PDF using available backend (Excel or LibreOffice)."""

    excel = wb = None
    try:
        import win32com.client  # type: ignore

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
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

    if success:
        return True

    try:
        outdir = os.path.dirname(pdf_path)
        subprocess.run(
            [
                "soffice",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                outdir,
                xlsx_path,
            ],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        return os.path.exists(pdf_path)
    except Exception:
        return False
