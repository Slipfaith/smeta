import hashlib
import sys
import time
import tempfile
import shutil
from pathlib import Path
from typing import Optional

import requests
from PySide6.QtWidgets import QApplication, QMessageBox

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


def check_for_updates(parent: Optional[object] = None, force: bool = False) -> None:
    """Check GitHub for a new release and update if confirmed by the user."""
    if not _should_check(force):
        return
    _save_check_time()
    lang = getattr(parent, "gui_lang", "ru")
    try:
        url = f"https://api.github.com/repos/{REPO}/releases/latest"
        data = requests.get(url, timeout=10).json()
        tag = data.get("tag_name", "")
        if tag.startswith("v"):
            tag = tag[1:]
        if tag <= APP_VERSION:
            if force:
                QMessageBox.information(parent, tr("Обновление", lang), tr("Установлена последняя версия", lang))
            return
        assets = {a["name"]: a["browser_download_url"] for a in data.get("assets", [])}
        exe_url = next((u for n, u in assets.items() if n.endswith(".exe")), None)
        sha_url = next((u for n, u in assets.items() if n.endswith(".sha256")), None)
        if not exe_url or not sha_url:
            return
        tmp_dir = Path(tempfile.mkdtemp(prefix="smeta_update"))
        exe_file = tmp_dir / "smeta.exe"
        sha_file = tmp_dir / "smeta.exe.sha256"
        _download(exe_url, exe_file)
        _download(sha_url, sha_file)
        expected = sha_file.read_text().strip().split()[0]
        actual = _sha256(exe_file)
        if expected != actual:
            QMessageBox.warning(parent, tr("Обновление", lang), tr("Ошибка проверки подлинности", lang))
            return
        reply = QMessageBox.question(
            parent,
            tr("Доступно обновление", lang),
            tr("Новая версия {0}. Закрыть программу для установки?", lang).format(tag),
        )
        if reply == QMessageBox.Yes:
            exe_path = Path(sys.argv[0]).resolve()
            old_path = exe_path.with_name("old.exe")
            try:
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
            except Exception as e:  # pragma: no cover - safety
                QMessageBox.critical(
                    parent,
                    tr("Обновление", lang),
                    tr("Ошибка при установке обновления: {0}", lang).format(e),
                )
    except Exception as e:  # pragma: no cover - network errors etc.
        if force:
            QMessageBox.warning(parent, tr("Обновление", lang), str(e))
