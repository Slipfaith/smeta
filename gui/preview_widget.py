# gui/preview_widget.py
import os
import tempfile
import subprocess
from typing import Dict, Any, Optional

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel
from PySide6.QtPdfWidgets import QPdfView
from PySide6.QtPdf import QPdfDocument

from logic.excel_exporter import ExcelExporter
from logic.legal_entities import load_legal_entities


class PreviewWidget(QWidget):
    """Правая панель предпросмотра: собирает XLSX и показывает PDF-рендер шаблона."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.viewer = QPdfView()
        self.doc = QPdfDocument(self)
        self.viewer.setDocument(self.doc)
        self.viewer.setPageMode(QPdfView.PageMode.MultiPage)
        self.status = QLabel("Предпросмотр: нет данных")
        self.btn_refresh = QPushButton("Обновить")
        self.btn_open_xlsx = QPushButton("Открыть XLSX")
        self._last_xlsx: Optional[str] = None
        self._last_pdf: Optional[str] = None

        self.btn_refresh.clicked.connect(self._refresh_button)
        self.btn_open_xlsx.clicked.connect(self._open_xlsx)
        self.legal_entities = load_legal_entities()

        top = QHBoxLayout()
        top.addWidget(self.status)
        top.addStretch(1)
        top.addWidget(self.btn_open_xlsx)
        top.addWidget(self.btn_refresh)

        lay = QVBoxLayout()
        lay.addLayout(top)
        lay.addWidget(self.viewer, 1)
        self.setLayout(lay)

        self._debounce = QTimer(self)
        self._debounce.setSingleShot(True)
        self._debounce.setInterval(500)  # 0.5s
        self._debounce.timeout.connect(self._do_render)
        self._pending_data: Optional[Dict[str, Any]] = None

    # публичное API
    def render_later(self, project_data: Dict[str, Any]) -> None:
        """Запланировать обновление превью (дебаунс)."""
        self._pending_data = project_data
        self._debounce.start()

    # кнопки
    def _refresh_button(self):
        if self._pending_data is not None:
            self._debounce.stop()
            self._do_render()

    def _open_xlsx(self):
        if self._last_xlsx and os.path.exists(self._last_xlsx):
            if os.name == "nt":
                os.startfile(self._last_xlsx)  # type: ignore[attr-defined]
            else:
                subprocess.Popen(["xdg-open", self._last_xlsx])

    # основная логика
    def _do_render(self):
        data = self._pending_data or {}
        self.status.setText("Собираю превью…")
        QApplication = None  # just to keep linter quiet
        try:
            tmpdir = tempfile.mkdtemp(prefix="pcalc_")
            xlsx = os.path.join(tmpdir, "preview.xlsx")
            pdf = os.path.join(tmpdir, "preview.pdf")

            entity_name = data.get("legal_entity")
            template_path = self.legal_entities.get(entity_name)
            ok = ExcelExporter(template_path, currency=data.get("currency", "RUB")).export_to_excel(data, xlsx)
            if not ok:
                self.status.setText("Ошибка экспорта XLSX (см. консоль).")
                return

            if not self._xlsx_to_pdf(xlsx, pdf):
                self.status.setText("Не удалось конвертировать в PDF (нужен Excel или LibreOffice).")
                self.doc.load("")
                return

            self.doc.load(pdf)
            self.viewer.setZoomMode(QPdfView.ZoomMode.FitToWidth)
            self._last_xlsx = xlsx
            self._last_pdf = pdf
            self.status.setText("Предпросмотр обновлён.")
        except Exception as e:
            self.status.setText(f"Ошибка превью: {e}")

    def _xlsx_to_pdf(self, xlsx_path: str, pdf_path: str) -> bool:
        # 1) Пытаемся через установленный Excel (COM)
        try:
            import win32com.client  # type: ignore
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            wb = excel.Workbooks.Open(xlsx_path)
            # 0 = xlTypePDF
            wb.ExportAsFixedFormat(0, pdf_path)
            wb.Close(False)
            excel.Quit()
            return os.path.exists(pdf_path)
        except Exception:
            pass

        # 2) Пытаемся через LibreOffice
        try:
            outdir = os.path.dirname(pdf_path)
            subprocess.run(
                ["soffice", "--headless", "--convert-to", "pdf", "--outdir", outdir, xlsx_path],
                check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
            )
            return os.path.exists(pdf_path)
        except Exception:
            return False
