"""Entry point for the Translation Cost Calculator application."""

from __future__ import annotations

import argparse
from pathlib import Path
import sys

from resource_utils import resource_path

from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon
from PySide6.QtCore import QTimer

from gui.main_window import TranslationCostCalculator
from logic.excel_process import close_excel_processes
from logic.logging_utils import setup_logging
from logic.activity_logger import log_user_action


def _parse_cli_arguments(argv: list[str]) -> tuple[Path | None, list[str]]:
    """Extract arguments intended for this module and leave the rest for Qt.

    Parameters
    ----------
    argv:
        Raw command-line arguments.

    Returns
    -------
    tuple[Path | None, list[str]]
        A tuple containing the path that should be used for the ready-signal
        handshake (if any) and the remaining arguments that should be passed to
        :class:`QApplication`.
    """

    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--ready-signal", dest="ready_signal", default=None)

    parsed, remaining = parser.parse_known_args(argv[1:])
    ready_signal = Path(parsed.ready_signal) if parsed.ready_signal else None

    # Ensure Qt receives only the arguments that are relevant to it.
    qt_arguments = [argv[0], *remaining]
    return ready_signal, qt_arguments


class _ReadySignal:
    """Handle coordination with the splash screen launcher."""

    def __init__(self, target: Path | None) -> None:
        self._target = target
        if self._target is None:
            return

        try:
            self._target.parent.mkdir(parents=True, exist_ok=True)
            if self._target.exists():
                self._target.unlink()
        except OSError:
            # The splash screen is best-effort; failures should not prevent the
            # main window from starting up.
            self._target = None

    def notify_ready(self) -> None:
        if self._target is None:
            return

        try:
            self._target.touch(exist_ok=True)
        except OSError:
            # Silently ignore failures to keep the application resilient.
            pass


def main() -> int:
    ready_signal_path, qt_arguments = _parse_cli_arguments(sys.argv)
    ready_signal = _ReadySignal(ready_signal_path)

    log_path = setup_logging()
    log_user_action(
        "Запуск приложения",
        details={"Файл журнала": str(log_path)},
        snapshot=None,
    )

    app = QApplication(qt_arguments)

    icon_path = resource_path("rateapp.ico")
    icon = QIcon(str(icon_path))
    app.setWindowIcon(icon)

    window = TranslationCostCalculator()
    window.setWindowIcon(icon)
    window.show()

    QTimer.singleShot(0, ready_signal.notify_ready)
    return app.exec()


if __name__ == "__main__":
    try:
        sys.exit(main())
    finally:
        # Ensure no hanging Excel processes remain if the application crashes.
        close_excel_processes()
