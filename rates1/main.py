# main.py
import os
import sys

from PySide6.QtWidgets import QApplication

if __package__ in (None, ""):
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(current_dir)
    for path in (current_dir, project_root):
        if path not in sys.path:
            sys.path.append(path)

    from app import RateApp
else:
    from .app import RateApp

from utils import apply_theme

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    apply_theme(app)

    window = RateApp()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
