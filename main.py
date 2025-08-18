import sys
from PySide6.QtWidgets import QApplication
from gui.main_window import TranslationCostCalculator

def main():
    app = QApplication(sys.argv)
    window = TranslationCostCalculator()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
