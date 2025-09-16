import contextlib
import hashlib
import os
import shutil
import subprocess
import sys
import tempfile
import textwrap
import time
from pathlib import Path
from typing import Optional

import requests
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QMessageBox, QProgressDialog

from logic.translation_config import tr
from .version import APP_VERSION

CHECK_INTERVAL = 24 * 60 * 60  # 24 hours
REPO = "Slipfaith/smeta"
LAST_CHECK_FILE = Path(tempfile.gettempdir()) / "smeta_last_update_check"


def _should_check(force: bool) -> bool:
    if force:
        return True
    if not LAST_CHECK_FILE.exists():
        return True
    return time.time() - LAST_CHECK_FILE.stat().st_mtime > CHECK_INTERVAL


def _save_check_time() -> None:
    LAST_CHECK_FILE.touch()


def _download(url: str, dest: Path) -> None:
    resp = requests.get(url, timeout=20)
    resp.raise_for_status()
    with open(dest, "wb") as f:
        f.write(resp.content)


def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


def _escape_for_batch(path: Path) -> str:
    """Escape characters that have special meaning in batch scripts."""

    return str(path).replace("%", "%%")


def _schedule_windows_install(exe_path: Path, downloaded_exe: Path, old_path: Path) -> None:
    """Prepare and launch a helper script that installs the update on Windows."""

    update_path = exe_path.with_name(f"{exe_path.stem}_update{exe_path.suffix}")
    script_path = exe_path.with_name("smeta_update.bat")
    if script_path.exists():
        script_path.unlink()
    script_template = textwrap.dedent(
        """\
        @echo off
        setlocal enableextensions
        set "TARGET={target}"
        set "NEW={new}"
        set "OLD={old}"
        set "SCRIPT=%~f0"
        if exist "%OLD%" del /f /q "%OLD%" >nul 2>&1
        :wait
        move /Y "%TARGET%" "%OLD%" >nul 2>&1
        if errorlevel 1 (
            timeout /t 1 /nobreak >nul
            goto wait
        )
        move /Y "%NEW%" "%TARGET%" >nul 2>&1
        start "" "%TARGET%"
        del /f /q "%OLD%" >nul 2>&1
        del /f /q "%SCRIPT%" >nul 2>&1
        """
    ).format(
        target=_escape_for_batch(exe_path),
        new=_escape_for_batch(update_path),
        old=_escape_for_batch(old_path),
    )
    script_path.write_text(script_template, encoding="utf-8")
    if update_path.exists():
        update_path.unlink()
    try:
        shutil.move(str(downloaded_exe), str(update_path))
    except Exception:
        with contextlib.suppress(Exception):
            script_path.unlink()
        raise
    creationflags = 0
    if hasattr(subprocess, "CREATE_NO_WINDOW"):
        creationflags |= subprocess.CREATE_NO_WINDOW
    try:
        subprocess.Popen(["cmd", "/c", str(script_path)], creationflags=creationflags)
    except Exception:
        with contextlib.suppress(Exception):
            shutil.move(str(update_path), str(downloaded_exe))
            script_path.unlink()
        raise


def check_for_updates(parent: Optional[object] = None, force: bool = False) -> None:
    """Check GitHub for a new release and update if confirmed by the user."""
    if not _should_check(force):
        return
    _save_check_time()
    lang = getattr(parent, "gui_lang", "ru")
    progress: Optional[QProgressDialog] = None

    def _update_progress(message: str) -> None:
        if progress is None:
            return
        progress.setLabelText(message)
        QApplication.processEvents()

    try:
        if parent is not None:
            progress = QProgressDialog(parent)
            progress.setWindowTitle(tr("Обновление", lang))
            progress.setLabelText(tr("Проверка обновления...", lang))
            progress.setCancelButton(None)
            progress.setWindowModality(Qt.WindowModal)
            progress.setMinimumDuration(0)
            progress.setRange(0, 0)
            progress.setAutoClose(False)
            progress.setAutoReset(False)
            progress.show()
            QApplication.processEvents()
        url = f"https://api.github.com/repos/{REPO}/releases/latest"
        data = requests.get(url, timeout=10).json()
        tag = data.get("tag_name", "")
        if tag.startswith("v"):
            tag = tag[1:]
        if tag <= APP_VERSION:
            if progress is not None:
                progress.close()
            if force:
                QMessageBox.information(parent, tr("Обновление", lang), tr("Установлена последняя версия", lang))
            return
        assets = {a["name"]: a["browser_download_url"] for a in data.get("assets", [])}
        exe_url = next((u for n, u in assets.items() if n.endswith(".exe")), None)
        sha_url = next((u for n, u in assets.items() if n.endswith(".sha256")), None)
        if not exe_url or not sha_url:
            if progress is not None:
                progress.close()
            return
        tmp_dir = Path(tempfile.mkdtemp(prefix="smeta_update"))
        exe_file = tmp_dir / "smeta.exe"
        sha_file = tmp_dir / "smeta.exe.sha256"
        try:
            _update_progress(tr("Загрузка обновления...", lang))
            _download(exe_url, exe_file)
            _download(sha_url, sha_file)
            _update_progress(tr("Проверка файла обновления...", lang))
            expected = sha_file.read_text().strip().split()[0]
            actual = _sha256(exe_file)
            if expected != actual:
                if progress is not None:
                    progress.close()
                QMessageBox.warning(parent, tr("Обновление", lang), tr("Ошибка проверки подлинности", lang))
                return
            if progress is not None:
                progress.close()
            reply = QMessageBox.question(
                parent,
                tr("Доступно обновление", lang),
                tr("Новая версия {0}. Закрыть программу для установки?", lang).format(tag),
            )
            if reply == QMessageBox.Yes:
                exe_path = Path(sys.argv[0]).resolve()
                old_path = exe_path.with_name("old.exe")
                try:
                    if os.name == "nt" and exe_path.suffix.lower() == ".exe":
                        try:
                            _schedule_windows_install(exe_path, exe_file, old_path)
                        except Exception as install_error:  # pragma: no cover - safety
                            QMessageBox.critical(
                                parent,
                                tr("Обновление", lang),
                                tr("Не удалось подготовить установку обновления: {0}", lang).format(install_error),
                            )
                            return
                        QMessageBox.information(
                            parent,
                            tr("Обновление", lang),
                            tr("Обновление скачано. Программа закроется для установки обновления.", lang),
                        )
                        QApplication.quit()
                        return
                    if old_path.exists():
                        old_path.unlink()
                    exe_path.rename(old_path)
                    shutil.move(str(exe_file), str(exe_path))
                    QMessageBox.information(
                        parent,
                        tr("Обновление", lang),
                        tr("Обновление установлено. Перезапустите программу.", lang),
                    )
                    QApplication.quit()
                    return
                except Exception as e:  # pragma: no cover - safety
                    QMessageBox.critical(
                        parent,
                        tr("Обновление", lang),
                        tr("Ошибка при установке обновления: {0}", lang).format(e),
                    )
        finally:
            with contextlib.suppress(FileNotFoundError):
                sha_file.unlink()
            shutil.rmtree(tmp_dir, ignore_errors=True)
    except Exception as e:  # pragma: no cover - network errors etc.
        if progress is not None:
            progress.close()
        if force:
            QMessageBox.warning(parent, tr("Обновление", lang), str(e))
    finally:
        if progress is not None:
            progress.close()
            QApplication.processEvents()
