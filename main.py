import sys
from PySide6.QtWidgets import QApplication
from gui.main_window import TranslationCostCalculator
from logic.excel_process import close_excel_processes

def main() -> int:
    app = QApplication(sys.argv)
    window = TranslationCostCalculator()
    window.show()
    return app.exec()


if __name__ == "__main__":
    try:
        sys.exit(main())
    finally:
        # Ensure no hanging Excel processes remain if the application crashes.
        close_excel_processes()
