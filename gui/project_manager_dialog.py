from typing import List, Dict

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QLineEdit, QPushButton, QComboBox, QMessageBox
)

from logic.translation_config import tr


class ProjectManagerDialog(QDialog):
    """Диалог для выбора и сохранения проджект-менеджеров."""

    def __init__(
        self,
        managers: List[Dict[str, str]],
        current_index: int = -1,
        lang: str = "ru",
        parent=None,
    ):
        super().__init__(parent)
        self.lang = lang
        self.setWindowTitle(tr("Проджект менеджер", lang))
        self.managers = managers
        self.current_index = current_index

        layout = QVBoxLayout()

        self.combo = QComboBox()
        self.combo.addItem(tr("<Новый менеджер>", lang))
        for m in self.managers:
            display = f"{m.get('name_ru') or m.get('name_en')} ({m.get('email')})"
            self.combo.addItem(display)
        if 0 <= current_index < len(self.managers):
            self.combo.setCurrentIndex(current_index + 1)
        self.combo.currentIndexChanged.connect(self.on_combo_changed)
        layout.addWidget(self.combo)

        layout.addWidget(QLabel(tr("Имя и фамилия (RU)", lang) + ":"))
        self.name_ru_edit = QLineEdit()
        layout.addWidget(self.name_ru_edit)

        layout.addWidget(QLabel(tr("Имя и фамилия (EN)", lang) + ":"))
        self.name_en_edit = QLineEdit()
        layout.addWidget(self.name_en_edit)

        layout.addWidget(QLabel(tr("Email", lang) + ":"))
        self.email_edit = QLineEdit()
        layout.addWidget(self.email_edit)

        save_btn = QPushButton(tr("Сохранить", lang))
        save_btn.clicked.connect(self.accept_dialog)
        layout.addWidget(save_btn)

        self.setLayout(layout)
        self.on_combo_changed(self.combo.currentIndex())

    def on_combo_changed(self, index: int) -> None:
        if index <= 0:
            self.name_ru_edit.clear()
            self.name_en_edit.clear()
            self.email_edit.clear()
        else:
            m = self.managers[index - 1]
            self.name_ru_edit.setText(m.get("name_ru", ""))
            self.name_en_edit.setText(m.get("name_en", ""))
            self.email_edit.setText(m.get("email", ""))

    def accept_dialog(self) -> None:
        name_ru = self.name_ru_edit.text().strip()
        name_en = self.name_en_edit.text().strip()
        email = self.email_edit.text().strip()
        if not (name_ru and name_en and email):
            QMessageBox.warning(
                self, tr("Ошибка", self.lang), tr("Заполните все поля", self.lang)
            )
            return
        data = {"name_ru": name_ru, "name_en": name_en, "email": email}
        if self.combo.currentIndex() <= 0:
            self.managers.append(data)
            self.current_index = len(self.managers) - 1
        else:
            self.managers[self.combo.currentIndex() - 1] = data
            self.current_index = self.combo.currentIndex() - 1
        self.accept()

    def result(self):
        return self.managers, self.current_index
