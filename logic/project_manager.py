import os
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Optional, TYPE_CHECKING

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QFileDialog, QMessageBox

from logic.excel_exporter import ExcelExporter
from logic.pdf_exporter import xlsx_to_pdf
from logic.progress import Progress
from logic.project_data import ProjectData
from logic.project_io import (
    load_project as load_project_file,
    save_project as save_project_file,
)
from logic.translation_config import tr

if TYPE_CHECKING:  # pragma: no cover - for type checkers only
    from gui.main_window import TranslationCostCalculator


class ProjectManager:
    """Encapsulates project-related operations for the main window."""

    def __init__(self, window: "TranslationCostCalculator") -> None:
        self.window = window

    # ------------------------------------------------------------------
    # Data preparation
    # ------------------------------------------------------------------
    def collect_project_data(self) -> ProjectData:
        return ProjectData.from_window(self.window)

    # ------------------------------------------------------------------
    # Persistence
    # ------------------------------------------------------------------
    def save_project(self) -> None:
        window = self.window
        if not window.project_name_edit.text().strip():
            QMessageBox.warning(window, "Ошибка", "Введите название проекта")
            return
        project_data = self.collect_project_data()
        filename = f"Проект_{project_data.project_slug}.json"
        file_path = self._ask_save_file(
            caption="Сохранить проект",
            filename=filename,
            filter_pattern="JSON files (*.json)",
        )
        if not file_path:
            return
        if save_project_file(project_data.to_mapping(), str(file_path)):
            QMessageBox.information(window, "Успех", f"Проект сохранен: {file_path}")
        else:
            QMessageBox.critical(window, "Ошибка", "Не удалось сохранить проект")

    def load_project(self) -> None:
        window = self.window
        file_path = self._ask_open_file(
            caption="Загрузить проект", filter_pattern="JSON files (*.json)"
        )
        if not file_path:
            return
        project_data = load_project_file(str(file_path))
        if project_data is None:
            QMessageBox.critical(window, "Ошибка", "Не удалось загрузить проект")
            return
        window.load_project_data(project_data)
        QMessageBox.information(window, "Успех", "Проект загружен")

    # ------------------------------------------------------------------
    # Export
    # ------------------------------------------------------------------
    def save_excel(self) -> None:
        window = self.window
        export_lang = "ru" if window.lang_display_ru else "en"
        project_data = self._prepare_export_data(
            require_project_name=False, require_services=True
        )
        if project_data is None:
            return

        currency = project_data.currency
        date_str = datetime.now().strftime("%Y-%m-%d")
        filename = (
            f"{date_str}-{project_data.legal_entity_slug}-{currency}-{project_data.client_slug}.xlsx"
        )

        file_path = self._ask_save_file(
            caption="Сохранить Excel файл",
            filename=filename,
            filter_pattern="Excel files (*.xlsx)",
        )
        if not file_path:
            return

        template_path = window.legal_entities.get(project_data.legal_entity)
        exporter = ExcelExporter(
            template_path,
            currency=currency,
            lang=export_lang,
        )
        with Progress(parent=window) as progress:
            success = exporter.export_to_excel(
                project_data.to_mapping(),
                str(file_path),
                progress_callback=progress.on_progress,
            )
        if success:
            self._schedule_file_saved_message(file_path)
        else:
            QMessageBox.critical(window, "Ошибка", "Не удалось сохранить файл")

    def save_pdf(self) -> None:
        window = self.window
        export_lang = "ru" if window.lang_display_ru else "en"
        project_data = self._prepare_export_data(
            require_project_name=True, require_services=False
        )
        if project_data is None:
            return

        currency = project_data.currency
        date_str = datetime.now().strftime("%Y-%m-%d")
        filename = (
            f"{date_str}-{project_data.legal_entity_slug}-{currency}-{project_data.client_slug}.pdf"
        )
        file_path = self._ask_save_file(
            caption="Сохранить PDF файл",
            filename=filename,
            filter_pattern="PDF files (*.pdf)",
        )
        if not file_path:
            return
        template_path = window.legal_entities.get(project_data.legal_entity)
        exporter = ExcelExporter(
            template_path,
            currency=currency,
            lang=export_lang,
        )
        with Progress(parent=window) as progress:
            def on_excel_progress(percent: int, message: str) -> None:
                progress.on_progress(int(percent * 0.8), message)

            try:
                with tempfile.TemporaryDirectory() as tmpdir:
                    tmp_dir_path = Path(tmpdir)
                    xlsx_path = tmp_dir_path / "quotation.xlsx"
                    pdf_path = tmp_dir_path / "quotation.pdf"
                    if not exporter.export_to_excel(
                        project_data.to_mapping(),
                        str(xlsx_path),
                        fit_to_page=True,
                        progress_callback=on_excel_progress,
                    ):
                        QMessageBox.critical(window, "Ошибка", "Не удалось подготовить файл")
                        return
                    progress.set_label("Конвертация в PDF")
                    progress.set_value(80)
                    if not xlsx_to_pdf(str(xlsx_path), str(pdf_path), lang=export_lang):
                        QMessageBox.critical(
                            window, "Ошибка", "Не удалось конвертировать в PDF"
                        )
                        return
                    progress.set_value(100)
                    shutil.copyfile(str(pdf_path), str(file_path))
                self._schedule_file_saved_message(file_path)
            except Exception as exc:  # pragma: no cover - defensive
                QMessageBox.critical(
                    window, "Ошибка", f"Не удалось сохранить PDF: {exc}"
                )

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def _show_file_saved_message(self, file_path: Path) -> None:
        window = self.window
        file_path = Path(file_path)
        msg_box = QMessageBox(window)
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setWindowTitle("Успех")
        msg_box.setText(f"Файл сохранен:\n{file_path}")
        open_button = msg_box.addButton("Открыть папку", QMessageBox.ActionRole)
        msg_box.addButton(QMessageBox.Ok)
        msg_box.setDefaultButton(QMessageBox.Ok)
        msg_box.exec()
        if msg_box.clickedButton() == open_button:
            self._reveal_file_in_explorer(file_path)

    def _schedule_file_saved_message(self, file_path: Path) -> None:
        """Show the "file saved" dialog on the next event loop iteration.

        Export operations run synchronously and keep the event loop busy until
        they finish, which could postpone the appearance of the success
        message even though the file is already written to disk.  Scheduling
        the dialog with a zero-timeout QTimer ensures that it pops up as soon
        as the export task releases control back to the UI thread, keeping the
        feedback immediate for the user.
        """
        QTimer.singleShot(
            0, lambda path=Path(file_path): self._show_file_saved_message(path)
        )

    def _reveal_file_in_explorer(self, file_path: Path) -> None:
        window = self.window
        file_path = Path(file_path)
        if not file_path.exists():
            QMessageBox.warning(
                window, "Ошибка", "Файл не найден для открытия в проводнике"
            )
            return
        try:
            if sys.platform.startswith("win"):
                subprocess.run(
                    ["explorer", f"/select,{file_path.resolve()}"], check=False
                )
            elif sys.platform == "darwin":
                subprocess.run(["open", "-R", str(file_path)], check=False)
            else:
                directory = file_path.parent if file_path.parent != Path() else Path(".")
                subprocess.run(["xdg-open", str(directory)], check=False)
        except Exception as exc:  # pragma: no cover - defensive
            QMessageBox.warning(
                window,
                "Ошибка",
                f"Не удалось открыть проводник:\n{exc}",
            )

    def _ask_save_file(
        self, *, caption: str, filename: str, filter_pattern: str
    ) -> Optional[Path]:
        file_path, _ = QFileDialog.getSaveFileName(
            self.window, caption, filename, filter_pattern
        )
        return Path(file_path) if file_path else None

    def _ask_open_file(
        self, *, caption: str, filter_pattern: str
    ) -> Optional[Path]:
        file_path, _ = QFileDialog.getOpenFileName(
            self.window, caption, "", filter_pattern
        )
        return Path(file_path) if file_path else None

    def _prepare_export_data(
        self, *, require_project_name: bool, require_services: bool
    ) -> Optional[ProjectData]:
        window = self.window
        lang = window.gui_lang

        if not window.get_selected_legal_entity():
            QMessageBox.warning(
                window,
                tr("Предупреждение", lang),
                tr("Выберите юрлицо", lang),
            )
            return None

        if require_project_name and not window.project_name_edit.text().strip():
            QMessageBox.warning(window, "Ошибка", "Введите название проекта")
            return None

        if not window.client_name_edit.text().strip():
            QMessageBox.warning(window, "Ошибка", "Введите название клиента")
            return None

        currency = window.get_current_currency_code()
        if not currency:
            QMessageBox.warning(
                window,
                tr("Предупреждение", lang),
                tr("Выберите валюту", lang),
            )
            return None

        project_data = self.collect_project_data()
        if require_services and not project_data.has_any_services():
            QMessageBox.warning(window, "Ошибка", "Добавьте хотя бы одну услугу")
            return None

        if project_data.has_zero_setup_rates():
            message = tr("Ставка для \"{0}\" равна 0. Продолжить?", lang).format(
                tr("Запуск и управление проектом", lang)
            )
            reply = QMessageBox.question(
                window,
                tr("Предупреждение", lang),
                message,
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply == QMessageBox.No:
                return None

        return project_data
