# logic/pdf_exporter.py
import os
import logging
import tempfile
import gc
from typing import Dict, Any

from .excel_exporter import ExcelExporter
from .com_utils import get_excel_app, close_excel_app

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
            if not exporter.export_to_excel(
                project_data,
                xlsx_path,
                fit_to_page=True,
                restore_images=False,
            ):
                return False
            if not xlsx_to_pdf(xlsx_path, output_path, template_path, lang=lang):
                raise RuntimeError("Не удалось конвертировать в PDF")
        return True
    except Exception as e:
        logger.exception("PDF export failed")
        print(f"[PdfExporter] Ошибка экспорта PDF: {e}")
        return False


def xlsx_to_pdf(
    xlsx_path: str, pdf_path: str, template_path: str, lang: str = "ru"
) -> bool:
    """Convert an XLSX file to PDF using a single Excel session.

    The function restores images from ``template_path``, adjusts page setup,
    toggles heavy Excel features off for faster processing and finally exports
    the workbook to ``pdf_path``.  All changes to the Excel environment are
    reverted in a ``finally`` block.
    """

    excel = tpl_wb = out_wb = None
    orig_decimal = orig_thousands = orig_use_sys = None
    orig_screen = orig_events = orig_alerts = orig_calc = None
    custom_sep = None
    try:
        excel = get_excel_app()

        # store and tweak performance toggles
        try:
            orig_screen = excel.ScreenUpdating
            orig_events = excel.EnableEvents
            orig_alerts = excel.DisplayAlerts
            orig_calc = excel.Calculation
            excel.ScreenUpdating = False
            excel.EnableEvents = False
            excel.DisplayAlerts = False
            # xlCalculationManual = -4135
            excel.Calculation = -4135
        except Exception:
            pass

        lang_lc = lang.lower()
        if lang_lc.startswith("en"):
            custom_sep = (".", ",")
        elif lang_lc.startswith("ru"):
            custom_sep = (",", " ")

        if custom_sep is not None:
            try:
                orig_decimal = excel.DecimalSeparator
                orig_thousands = excel.ThousandsSeparator
                orig_use_sys = excel.UseSystemSeparators
                excel.DecimalSeparator, excel.ThousandsSeparator = custom_sep
                excel.UseSystemSeparators = False
            except Exception:
                pass

        tpl_wb = excel.Workbooks.Open(os.path.abspath(template_path))
        out_wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))
        tpl_ws = tpl_wb.Worksheets("Quotation")
        out_ws = out_wb.Worksheets("Quotation")

        # restore pictures from template
        while out_ws.Shapes.Count > 0:
            out_ws.Shapes(1).Delete()
        for shape in tpl_ws.Shapes:
            if int(getattr(shape, "Type", 0)) == 13:  # msoPicture
                shape.Copy()
                out_ws.Paste()
                pasted = out_ws.Shapes(out_ws.Shapes.Count)
                pasted.Left = shape.Left
                pasted.Top = shape.Top
                pasted.Width = shape.Width
                pasted.Height = shape.Height

        tpl_wb.Close(False)
        tpl_wb = None

        # PageSetup adjustments
        try:
            ps = out_ws.PageSetup
            ps.Zoom = False
            ps.FitToPagesTall = 1
            ps.FitToPagesWide = 1
        except Exception:
            pass

        out_wb.ExportAsFixedFormat(0, pdf_path)
        success = os.path.exists(pdf_path)
    except Exception:
        success = False
    finally:
        if out_wb is not None:
            try:
                out_wb.Close(False)
            except Exception:
                pass
            out_wb = None
        if tpl_wb is not None:
            try:
                tpl_wb.Close(False)
            except Exception:
                pass
            tpl_wb = None
        if excel is not None:
            try:
                if custom_sep is not None and orig_use_sys is not None:
                    try:
                        excel.DecimalSeparator = orig_decimal
                        excel.ThousandsSeparator = orig_thousands
                        excel.UseSystemSeparators = orig_use_sys
                    except Exception:
                        pass
                try:
                    if orig_screen is not None:
                        excel.ScreenUpdating = orig_screen
                    if orig_events is not None:
                        excel.EnableEvents = orig_events
                    if orig_alerts is not None:
                        excel.DisplayAlerts = orig_alerts
                    if orig_calc is not None:
                        excel.Calculation = orig_calc
                except Exception:
                    pass
                close_excel_app(excel)
            except Exception:
                pass
            excel = None
        gc.collect()

    return success
