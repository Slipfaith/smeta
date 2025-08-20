import json
import os
from typing import List, Dict, Tuple

from .user_config import get_appdata_dir

PM_FILE = "pm_history.json"

def _pm_path() -> str:
    return os.path.join(get_appdata_dir(), PM_FILE)

def load_pm_history() -> Tuple[List[Dict[str, str]], int]:
    path = _pm_path()
    if not os.path.exists(path):
        return [], -1
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
            managers = data.get("managers", []) if isinstance(data, dict) else []
            last = data.get("last_used", -1) if isinstance(data, dict) else -1
            if not isinstance(managers, list):
                managers = []
            return managers, last if isinstance(last, int) else -1
    except Exception:
        return [], -1

def save_pm_history(managers: List[Dict[str, str]], last_index: int) -> bool:
    path = _pm_path()
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump({"managers": managers, "last_used": last_index}, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False
