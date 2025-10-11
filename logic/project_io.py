import json
import logging

logger = logging.getLogger(__name__)

def save_project(data, path) -> bool:
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        logger.exception("Ошибка сохранения проекта")
        return False

def load_project(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        logger.exception("Ошибка загрузки проекта")
        return None
