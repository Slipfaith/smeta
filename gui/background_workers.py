"""Background workers for long-running GUI operations."""

from __future__ import annotations

import os
import shutil
import tempfile
import traceback
from typing import Dict, Any, List

from PySide6.QtCore import QObject, Signal, Slot

from logic.excel_exporter import ExcelExporter
from logic.pdf_exporter import xlsx_to_pdf
from logic.trados_xml_parser import parse_reports


class ParseReportsWorker(QObject):
    """Parse XML reports in a background thread."""

    finished = Signal(dict, list, dict)
    error = Signal(str, str)
    progress = Signal(int, str)

    def __init__(self, paths: List[str]) -> None:
        super().__init__()
        self._paths = paths

    @Slot()
    def run(self) -> None:
        try:
            self.progress.emit(0, "Анализ отчетов")
            data, warnings, report_sources = parse_reports(self._paths)
            self.progress.emit(100, "Готово")
            self.finished.emit(data, warnings, report_sources)
        except Exception as exc:  # noqa: BLE001 - propagate via signal
            self.error.emit(str(exc), traceback.format_exc())


class ExcelExportWorker(QObject):
    """Export project data to an Excel file in the background."""

    finished = Signal(str)
    error = Signal(str, str)
    progress = Signal(int, str)

    def __init__(
        self,
        project_data: Dict[str, Any],
        output_path: str,
        template_path: str,
        currency: str,
        lang: str,
        fit_to_page: bool = False,
    ) -> None:
        super().__init__()
        self._project_data = project_data
        self._output_path = output_path
        self._template_path = template_path
        self._currency = currency
        self._lang = lang
        self._fit_to_page = fit_to_page

    @Slot()
    def run(self) -> None:
        try:
            exporter = ExcelExporter(
                self._template_path,
                currency=self._currency,
                lang=self._lang,
            )
            success = exporter.export_to_excel(
                self._project_data,
                self._output_path,
                fit_to_page=self._fit_to_page,
                progress_callback=self._emit_progress,
            )
            if not success:
                raise RuntimeError("Не удалось сохранить файл")
            self.finished.emit(self._output_path)
        except Exception as exc:  # noqa: BLE001 - propagate via signal
            self.error.emit(str(exc), traceback.format_exc())

    def _emit_progress(self, percent: int, message: str) -> None:
        self.progress.emit(percent, message)


class PdfExportWorker(QObject):
    """Create a PDF export in the background."""

    finished = Signal(str)
    error = Signal(str, str)
    progress = Signal(int, str)

    def __init__(
        self,
        project_data: Dict[str, Any],
        output_path: str,
        template_path: str,
        currency: str,
        lang: str,
    ) -> None:
        super().__init__()
        self._project_data = project_data
        self._output_path = output_path
        self._template_path = template_path
        self._currency = currency
        self._lang = lang

    @Slot()
    def run(self) -> None:
        try:
            exporter = ExcelExporter(
                self._template_path,
                currency=self._currency,
                lang=self._lang,
            )
            with tempfile.TemporaryDirectory() as tmpdir:
                xlsx_path = os.path.join(tmpdir, "quotation.xlsx")
                pdf_path = os.path.join(tmpdir, "quotation.pdf")
                success = exporter.export_to_excel(
                    self._project_data,
                    xlsx_path,
                    fit_to_page=True,
                    progress_callback=self._emit_excel_progress,
                )
                if not success:
                    raise RuntimeError("Не удалось подготовить файл")
                self.progress.emit(80, "Конвертация в PDF")
                if not xlsx_to_pdf(xlsx_path, pdf_path, lang=self._lang):
                    raise RuntimeError("Не удалось конвертировать в PDF")
                shutil.copyfile(pdf_path, self._output_path)
            self.progress.emit(100, "Готово")
            self.finished.emit(self._output_path)
        except Exception as exc:  # noqa: BLE001 - propagate via signal
            self.error.emit(str(exc), traceback.format_exc())

    def _emit_excel_progress(self, percent: int, message: str) -> None:
        scaled = min(79, int(percent * 0.8))
        self.progress.emit(scaled, message)


class RatesImportWorker(QObject):
    """Load rates from Excel in a background thread."""

    finished = Signal(dict)
    error = Signal(str, str)

    def __init__(self, path: str, currency: str, rate_type: str) -> None:
        super().__init__()
        self._path = path
        self._currency = currency
        self._rate_type = rate_type

    @Slot()
    def run(self) -> None:
        try:
            from logic import rates_importer

            rates = rates_importer.load_rates_from_excel(
                self._path, self._currency, self._rate_type
            )
            self.finished.emit(rates)
        except Exception as exc:  # noqa: BLE001 - propagate via signal
            self.error.emit(str(exc), traceback.format_exc())
