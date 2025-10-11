import os
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from typing import Any, Dict, TYPE_CHECKING

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QFileDialog, QMessageBox

from logic.excel_exporter import ExcelExporter
from logic.pdf_exporter import xlsx_to_pdf
from logic.progress import Progress
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
    def collect_project_data(self) -> Dict[str, Any]:
        window = self.window
        data: Dict[str, Any] = {
            "project_name": window.project_name_edit.text(),
            "client_name": window.client_name_edit.text(),
            "contact_person": window.contact_person_edit.text(),
            "email": window.email_edit.text(),
            "legal_entity": window.get_selected_legal_entity(),
            "currency": window.get_current_currency_code(),
            "language_pairs": [],
            "additional_services": [],
            "pm_name": window.current_pm.get(
                "name_ru" if window.lang_display_ru else "name_en", ""
            ),
            "pm_email": window.current_pm.get("email", ""),
            "project_setup_fee": window.project_setup_fee_spin.value(),
            "project_setup_enabled": (
                window.project_setup_widget.is_enabled()
                if window.project_setup_widget
                else False
            ),
            "project_setup": (
                window.project_setup_widget.get_data()
                if window.project_setup_widget
                and window.project_setup_widget.is_enabled()
                else []
            ),
            "project_setup_discount_percent": (
                window.project_setup_widget.get_discount_percent()
                if window.project_setup_widget
                else 0.0
            ),
            "project_setup_markup_percent": (
                window.project_setup_widget.get_markup_percent()
                if window.project_setup_widget
                else 0.0
            ),
            "vat_rate": window.vat_spin.value() if window.vat_spin.isEnabled() else 0,
            "only_new_repeats_mode": window.only_new_repeats_mode,
        }

        total_discount_amount = 0.0
        total_markup_amount = 0.0
        project_setup_discount_amount = 0.0
        project_setup_markup_amount = 0.0

        if getattr(window, "project_setup_widget", None):
            project_setup_discount_amount = window.project_setup_widget.get_discount_amount()
            total_discount_amount += project_setup_discount_amount
            project_setup_markup_amount = window.project_setup_widget.get_markup_amount()
            total_markup_amount += project_setup_markup_amount

        data["project_setup_discount_amount"] = project_setup_discount_amount
        data["project_setup_markup_amount"] = project_setup_markup_amount

        for pair_key, pair_widget in window.language_pairs.items():
            pair_data = pair_widget.get_data()
            if pair_data["services"]:
                pair_data["header_title"] = window.pair_headers.get(
                    pair_key, pair_widget.pair_name
                )
                data["language_pairs"].append(pair_data)
                total_discount_amount += pair_data.get("discount_amount", 0.0)
                total_markup_amount += pair_data.get("markup_amount", 0.0)

        data["language_pairs"].sort(
            key=lambda x: (
                x.get("pair_name", "").split(" - ")[1]
                if " - " in x.get("pair_name", "")
                else x.get("pair_name", "")
            )
        )

        additional = window.additional_services_widget.get_data()
        if additional:
            data["additional_services"] = additional
            total_discount_amount += sum(
                block.get("discount_amount", 0.0) for block in additional
            )
            total_markup_amount += sum(
                block.get("markup_amount", 0.0) for block in additional
            )

        data["total_discount_amount"] = total_discount_amount
        data["total_markup_amount"] = total_markup_amount
        return data

    # ------------------------------------------------------------------
    # Persistence
    # ------------------------------------------------------------------
    def save_project(self) -> None:
        window = self.window
        if not window.project_name_edit.text().strip():
            QMessageBox.warning(window, "Ошибка", "Введите название проекта")
            return
        project_data = self.collect_project_data()
        project_name = project_data["project_name"].replace(" ", "_")
        filename = f"Проект_{project_name}.json"
        file_path, _ = QFileDialog.getSaveFileName(
            window, "Сохранить проект", filename, "JSON files (*.json)"
        )
        if not file_path:
            return
        if save_project_file(project_data, file_path):
            QMessageBox.information(window, "Успех", f"Проект сохранен: {file_path}")
        else:
            QMessageBox.critical(window, "Ошибка", "Не удалось сохранить проект")

    def load_project(self) -> None:
        window = self.window
        file_path, _ = QFileDialog.getOpenFileName(
            window, "Загрузить проект", "", "JSON files (*.json)"
        )
        if not file_path:
            return
        project_data = load_project_file(file_path)
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
        if not window.get_selected_legal_entity():
            lang = window.gui_lang
            QMessageBox.warning(
                window,
                tr("Предупреждение", lang),
                tr("Выберите юрлицо", lang),
            )
            return
        if not window.client_name_edit.text().strip():
            QMessageBox.warning(window, "Ошибка", "Введите название клиента")
            return
        project_data = self.collect_project_data()
        if (
            not project_data["language_pairs"]
            and not project_data["additional_services"]
        ):
            QMessageBox.warning(window, "Ошибка", "Добавьте хотя бы одну услугу")
            return

        if any(r.get("rate", 0) == 0 for r in project_data.get("project_setup", [])):
            lang = window.gui_lang
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
                return

        client_name = project_data["client_name"].replace(" ", "_")
        entity_for_file = window.get_selected_legal_entity().replace(" ", "_")
        currency = window.get_current_currency_code()
        date_str = datetime.now().strftime("%Y-%m-%d")
        filename = f"{date_str}-{entity_for_file}-{currency}-{client_name}.xlsx"

        file_path, _ = QFileDialog.getSaveFileName(
            window, "Сохранить Excel файл", filename, "Excel files (*.xlsx)"
        )
        if not file_path:
            return

        entity_name = window.get_selected_legal_entity()
        template_path = window.legal_entities.get(entity_name)
        exporter = ExcelExporter(
            template_path,
            currency=window.get_current_currency_code(),
            lang=export_lang,
        )
        with Progress(parent=window) as progress:
            success = exporter.export_to_excel(
                project_data, file_path, progress_callback=progress.on_progress
            )
        if success:
            self._schedule_file_saved_message(file_path)
        else:
            QMessageBox.critical(window, "Ошибка", "Не удалось сохранить файл")

    def save_pdf(self) -> None:
        window = self.window
        export_lang = "ru" if window.lang_display_ru else "en"
        if not window.get_selected_legal_entity():
            lang = window.gui_lang
            QMessageBox.warning(
                window,
                tr("Предупреждение", lang),
                tr("Выберите юрлицо", lang),
            )
            return
        if not window.project_name_edit.text().strip():
            QMessageBox.warning(window, "Ошибка", "Введите название проекта")
            return
        if not window.client_name_edit.text().strip():
            QMessageBox.warning(window, "Ошибка", "Введите название клиента")
            return
        project_data = self.collect_project_data()
        if any(r.get("rate", 0) == 0 for r in project_data.get("project_setup", [])):
            lang = window.gui_lang
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
                return
        client_name = project_data["client_name"].replace(" ", "_")
        entity_for_file = window.get_selected_legal_entity().replace(" ", "_")
        currency = window.get_current_currency_code()
        date_str = datetime.now().strftime("%Y-%m-%d")
        filename = f"{date_str}-{entity_for_file}-{currency}-{client_name}.pdf"
        file_path, _ = QFileDialog.getSaveFileName(
            window, "Сохранить PDF файл", filename, "PDF files (*.pdf)"
        )
        if not file_path:
            return
        template_path = window.legal_entities.get(window.get_selected_legal_entity())
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
                    xlsx_path = os.path.join(tmpdir, "quotation.xlsx")
                    pdf_path = os.path.join(tmpdir, "quotation.pdf")
                    if not exporter.export_to_excel(
                        project_data,
                        xlsx_path,
                        fit_to_page=True,
                        progress_callback=on_excel_progress,
                    ):
                        QMessageBox.critical(window, "Ошибка", "Не удалось подготовить файл")
                        return
                    progress.set_label("Конвертация в PDF")
                    progress.set_value(80)
                    if not xlsx_to_pdf(xlsx_path, pdf_path, lang=export_lang):
                        QMessageBox.critical(
                            window, "Ошибка", "Не удалось конвертировать в PDF"
                        )
                        return
                    progress.set_value(100)
                    shutil.copyfile(pdf_path, file_path)
                self._schedule_file_saved_message(file_path)
            except Exception as exc:  # pragma: no cover - defensive
                QMessageBox.critical(
                    window, "Ошибка", f"Не удалось сохранить PDF: {exc}"
                )

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def _show_file_saved_message(self, file_path: str) -> None:
        window = self.window
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

    def _schedule_file_saved_message(self, file_path: str) -> None:
        """Show the "file saved" dialog on the next event loop iteration.

        Export operations run synchronously and keep the event loop busy until
        they finish, which could postpone the appearance of the success
        message even though the file is already written to disk.  Scheduling
        the dialog with a zero-timeout QTimer ensures that it pops up as soon
        as the export task releases control back to the UI thread, keeping the
        feedback immediate for the user.
        """

        QTimer.singleShot(0, lambda path=file_path: self._show_file_saved_message(path))

    def _reveal_file_in_explorer(self, file_path: str) -> None:
        window = self.window
        if not os.path.exists(file_path):
            QMessageBox.warning(
                window, "Ошибка", "Файл не найден для открытия в проводнике"
            )
            return
        try:
            if sys.platform.startswith("win"):
                subprocess.run(
                    ["explorer", f"/select,{os.path.normpath(file_path)}"], check=False
                )
            elif sys.platform == "darwin":
                subprocess.run(["open", "-R", file_path], check=False)
            else:
                directory = os.path.dirname(file_path) or "."
                subprocess.run(["xdg-open", directory], check=False)
        except Exception as exc:  # pragma: no cover - defensive
            QMessageBox.warning(
                window,
                "Ошибка",
                f"Не удалось открыть проводник:\n{exc}",
            )
