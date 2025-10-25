import argparse
import sys

from resource_utils import resource_path

from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon

from gui.main_window import TranslationCostCalculator
from logic.excel_process import close_excel_processes
from logic.logging_utils import enable_console_logging, setup_logging
from logic.activity_logger import log_user_action


def _build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="RateApp launcher")
    parser.add_argument(
        "--console-log",
        action="store_true",
        help=(
            "stream log messages to the invoking console when possible; "
            "useful for debugging packaged builds"
        ),
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = _build_arg_parser()
    args = parser.parse_args(argv)

    if args.console_log:
        enable_console_logging()

    log_path = setup_logging()
    log_user_action(
        "Запуск приложения",
        details={"Файл журнала": str(log_path)},
        snapshot=None,
    )

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
