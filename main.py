import sys

from resource_utils import resource_path

from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon

from gui.main_window import TranslationCostCalculator
from logic.excel_process import close_excel_processes

def main() -> int:
    app = QApplication(sys.argv)

    icon_path = resource_path("rateapp.ico")
    icon = QIcon(str(icon_path))
    app.setWindowIcon(icon)

    window = TranslationCostCalculator()
    window.setWindowIcon(icon)
    window.show()
    return app.exec()


if __name__ == "__main__":
    try:
        sys.exit(main())
    finally:
        # Ensure no hanging Excel processes remain if the application crashes.
        close_excel_processes()
