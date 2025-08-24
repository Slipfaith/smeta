# gui/preview_widget.py
import os
import tempfile
import subprocess
from typing import Dict, Any, Optional

from PySide6.QtCore import Qt, QTimer, QThread, QObject, Signal
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QProgressDialog,
)
from PySide6.QtPdfWidgets import QPdfView
from PySide6.QtPdf import QPdfDocument

from logic.excel_exporter import ExcelExporter
from logic.legal_entities import load_legal_entities


class PreviewRenderWorker(QObject):
    """Генерирует превью в отдельном потоке."""

    finished = Signal(str, str)
    error = Signal(str)
    progress = Signal(int)

    def __init__(self, exporter: ExcelExporter, data: Dict[str, Any]):
        super().__init__()
        self._exporter = exporter
        self._data = data

    def run(self) -> None:
        try:
            tmpdir = tempfile.mkdtemp(prefix="pcalc_")
            xlsx = os.path.join(tmpdir, "preview.xlsx")
            pdf = os.path.join(tmpdir, "preview.pdf")

            self.progress.emit(0)
            if not self._exporter.export_to_excel(self._data, xlsx, fit_to_page=True):
                self.error.emit("Ошибка экспорта XLSX (см. консоль).")
                return

            self.progress.emit(50)
            if not self._exporter.xlsx_to_pdf(xlsx, pdf):
                self.error.emit(
                    "Не удалось конвертировать в PDF (нужен Excel или LibreOffice)."
                )
                return

            self.progress.emit(100)
            self.finished.emit(xlsx, pdf)
        except Exception as e:  # pragma: no cover - GUI code
            self.error.emit(f"Ошибка превью: {e}")


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
        self._pending_data = None
        self.status.setText("Собираю превью…")

        entity_name = data.get("legal_entity")
        template_path = self.legal_entities.get(entity_name)
        exporter = ExcelExporter(template_path, currency=data.get("currency", "RUB"))

        progress = QProgressDialog("Генерация превью...", "", 0, 100, self)
        progress.setWindowTitle("Пожалуйста, подождите")
        progress.setWindowModality(Qt.WindowModal)
        progress.setAutoClose(False)
        progress.setAutoReset(False)
        progress.setCancelButton(None)
        progress.show()

        worker = PreviewRenderWorker(exporter, data)
        thread = QThread(self)
        worker.moveToThread(thread)

        def cleanup() -> None:
            progress.close()
            thread.quit()
            thread.wait()
            thread.deleteLater()
            worker.deleteLater()

        def on_finished(xlsx: str, pdf: str) -> None:
            cleanup()
            self.doc.load(pdf)
            self.viewer.setZoomMode(QPdfView.ZoomMode.FitToWidth)
            self._last_xlsx = xlsx
            self._last_pdf = pdf
            self.status.setText("Предпросмотр обновлён.")

        def on_error(msg: str) -> None:
            cleanup()
            self.doc.load("")
            self.status.setText(msg)

        worker.finished.connect(on_finished)
        worker.error.connect(on_error)
        worker.progress.connect(progress.setValue)
        thread.started.connect(worker.run)
        thread.start()
