"""Standalone splash screen launcher for the Translation Cost Calculator."""

from __future__ import annotations

import argparse
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import Sequence

from PySide6.QtCore import QObject, Qt, QTimer, QUrl, Signal
from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QMessageBox
from PySide6.QtMultimedia import QAudioOutput, QMediaPlayer
from PySide6.QtMultimediaWidgets import QVideoWidget

from resource_utils import resource_path


class _SplashWindow(QWidget):
    """A frameless window that plays the provided animation on repeat."""

    def __init__(self, animation_path: Path) -> None:
        super().__init__()
        self.setWindowFlags(
            Qt.SplashScreen | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint
        )
        self.setAttribute(Qt.WA_TranslucentBackground, True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self._video_widget = QVideoWidget(self)
        layout.addWidget(self._video_widget)

        self._player = QMediaPlayer(self)
        self._audio_output = QAudioOutput(self)
        self._player.setAudioOutput(self._audio_output)
        self._player.setVideoOutput(self._video_widget)
        self._player.setSource(QUrl.fromLocalFile(str(animation_path)))
        self._player.mediaStatusChanged.connect(self._maybe_restart)
        self._player.play()

    def _maybe_restart(self, status: QMediaPlayer.MediaStatus) -> None:
        if status == QMediaPlayer.MediaStatus.EndOfMedia:
            self._player.setPosition(0)
            self._player.play()


class _ReadyWatcher(QObject):
    """Poll for the appearance of the ready-signal file."""

    ready = Signal()
    failed = Signal(int)

    def __init__(self, ready_path: Path, process: subprocess.Popen[bytes], interval: float) -> None:
        super().__init__()
        self._ready_path = ready_path
        self._process = process

        self._timer = QTimer(self)
        self._timer.setInterval(max(int(interval * 1000), 50))
        self._timer.timeout.connect(self._poll)
        self._timer.start()

    def _poll(self) -> None:
        if self._ready_path.exists():
            self._timer.stop()
            self.ready.emit()
            return

        exit_code = self._process.poll()
        if exit_code is not None:
            self._timer.stop()
            self.failed.emit(exit_code)


def _default_main_command() -> list[str]:
    rate_app = resource_path("RateApp.exe")
    if rate_app.exists():
        return [str(rate_app)]

    return [sys.executable, str(resource_path("main.py"))]


def _prepare_command(raw: Sequence[str], ready_path: Path) -> list[str]:
    command = list(raw) if raw else _default_main_command()

    if "--ready-signal" not in command:
        command = [*command, "--ready-signal", str(ready_path)]

    return command


def _parse_arguments(argv: Sequence[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Show a splash screen animation while launching the main application."
        )
    )
    parser.add_argument(
        "--animation",
        type=Path,
        default=resource_path("templates/splash.mov"),
        help="Path to the animation file (.mov) displayed in the splash window.",
    )
    parser.add_argument(
        "--timeout",
        type=float,
        default=60.0,
        help="Maximum time (in seconds) before the splash window closes automatically.",
    )
    parser.add_argument(
        "--poll-interval",
        type=float,
        default=0.2,
        help="How frequently (in seconds) to check if the main application is ready.",
    )
    parser.add_argument(
        "command",
        nargs=argparse.REMAINDER,
        help=(
            "Command used to start the main application. Prefix with '--' to "
            "separate it from splash arguments."
        ),
    )

    return parser.parse_args(argv[1:])


def _normalise_command(args: argparse.Namespace) -> list[str]:
    command = args.command
    if command and command[0] == "--":
        command = command[1:]
    return command


def _show_failure_message(exit_code: int) -> None:
    QMessageBox.critical(
        None,
        "Ошибка запуска",
        (
            "Не удалось запустить основное приложение. "
            f"Процесс завершился с кодом {exit_code}."
        ),
    )


def main(argv: Sequence[str] | None = None) -> int:
    argv = list(argv or sys.argv)
    args = _parse_arguments(argv)
    command = _normalise_command(args)

    animation_path = Path(args.animation)
    if not animation_path.exists():
        raise FileNotFoundError(
            f"Файл анимации '{animation_path}' не найден. Укажите корректный путь через --animation."
        )

    ready_path = Path(tempfile.gettempdir()) / (
        f"smeta_ready_{int(time.time() * 1000)}_{Path(argv[0]).stem}.flag"
    )
    if ready_path.exists():
        try:
            ready_path.unlink()
        except OSError:
            pass

    command_with_signal = _prepare_command(command, ready_path)
    process = subprocess.Popen(command_with_signal)

    app = QApplication([argv[0]])
    window = _SplashWindow(animation_path)
    window.show()

    watcher = _ReadyWatcher(ready_path, process, interval=args.poll_interval)
    watcher.ready.connect(app.quit)

    def _handle_failure(exit_code: int) -> None:
        _show_failure_message(exit_code)
        app.quit()

    watcher.failed.connect(_handle_failure)

    if args.timeout:
        QTimer.singleShot(int(args.timeout * 1000), app.quit)

    exit_code = app.exec()

    try:
        if ready_path.exists():
            ready_path.unlink()
    except OSError:
        pass

    return exit_code


if __name__ == "__main__":
    sys.exit(main())
