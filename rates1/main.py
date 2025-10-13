# main.py
import sys
from PySide6.QtWidgets import QApplication
from app import RateApp
from utils.dark_theme import apply_dark_theme
from utils.light_theme import apply_light_theme

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    apply_dark_theme(app)

    window = RateApp()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
