"""Runtime configurable service coefficients."""

from __future__ import annotations

from copy import deepcopy
import json
from pathlib import Path
from typing import Any, Dict, List, Tuple

from logic.user_config import get_appdata_dir


class ServiceConfig:
    """Конфигурация услуг и коэффициентов."""

    _CONFIG_FILE = "service_config.json"

    _DEFAULT_TRANSLATION_ROWS: List[Dict[str, Any]] = [
        {"key": "new", "name": "Перевод, новые слова (100%)", "multiplier": 1.0, "is_base": True},
        {"key": "fuzzy_75_94", "name": "Перевод, совпадения 75-94% (66%)", "multiplier": 0.66, "is_base": False},
        {"key": "fuzzy_95_99", "name": "Перевод, совпадения 95-99% (33%)", "multiplier": 0.33, "is_base": False},
        {"key": "reps_100_30", "name": "Перевод, повторы и 100% совпадения (30%)", "multiplier": 0.30, "is_base": False},
    ]

    _DEFAULT_ADDITIONAL_SERVICES: Dict[str, List[Dict[str, Any]]] = {
        "Верстка": [
            {"name": "InDesign верстка", "multiplier": 1.0, "is_base": True},
            {"name": "PowerPoint верстка", "multiplier": 1.0, "is_base": True},
            {"name": "PDF верстка", "multiplier": 1.0, "is_base": True},
            {"name": "Графика/Изображения", "multiplier": 1.0, "is_base": True},
        ],
        "Локализация мультимедиа": [
            {"name": "Создание субтитров", "multiplier": 1.0, "is_base": True},
            {"name": "Озвучка", "multiplier": 1.0, "is_base": True},
            {"name": "Видеомонтаж", "multiplier": 1.0, "is_base": True},
            {"name": "Синхронизация", "multiplier": 1.0, "is_base": True},
        ],
        "Тестирование/QA": [
            {"name": "Лингвистическое тестирование", "multiplier": 1.0, "is_base": True},
            {"name": "Функциональное тестирование", "multiplier": 1.0, "is_base": True},
            {"name": "Косметическое тестирование", "multiplier": 1.0, "is_base": True},
            {"name": "Финальная проверка", "multiplier": 1.0, "is_base": True},
        ],
        "Прочие услуги": [
            {"name": "Создание терминологии", "multiplier": 1.0, "is_base": True},
            {"name": "Подготовка Translation Memory", "multiplier": 1.0, "is_base": True},
            {"name": "Анализ CAT-инструмента", "multiplier": 1.0, "is_base": True},
            {"name": "Консультации", "multiplier": 1.0, "is_base": True},
        ],
    }

    TRANSLATION_ROWS: List[Dict[str, Any]] = []
    ROW_NAMES: List[str] = []
    ADDITIONAL_SERVICES: Dict[str, List[Dict[str, Any]]] = {}

    @classmethod
    def _config_path(cls) -> Path:
        return Path(get_appdata_dir()) / cls._CONFIG_FILE

    @classmethod
    def _default_config(cls) -> Dict[str, Any]:
        return {
            "translation_rows": deepcopy(cls._DEFAULT_TRANSLATION_ROWS),
            "additional_services": deepcopy(cls._DEFAULT_ADDITIONAL_SERVICES),
        }

    @classmethod
    def _load_raw_config(cls) -> Dict[str, Any]:
        path = cls._config_path()
        if not path.exists():
            return cls._default_config()
        try:
            with path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
            if isinstance(data, dict):
                return data
        except Exception:
            pass
        return cls._default_config()

    @classmethod
    def _validate_rows(
        cls, rows: Any, default_rows: List[Dict[str, Any]]
    ) -> List[Dict[str, Any]]:
        validated: List[Dict[str, Any]] = []
        if isinstance(rows, list):
            for row in rows:
                if not isinstance(row, dict):
                    continue
                name = str(row.get("name", "")).strip()
                if not name:
                    continue
                validated.append(
                    {
                        "key": row.get("key"),
                        "name": name,
                        "multiplier": float(row.get("multiplier", 1.0) or 0.0),
                        "is_base": bool(row.get("is_base", False)),
                    }
                )
        if not validated:
            return deepcopy(default_rows)
        if not any(r.get("is_base") for r in validated):
            validated[0]["is_base"] = True
        return validated

    @classmethod
    def _validate_additional(
        cls, services: Any, default_services: Dict[str, List[Dict[str, Any]]]
    ) -> Dict[str, List[Dict[str, Any]]]:
        if not isinstance(services, dict):
            return deepcopy(default_services)
        result: Dict[str, List[Dict[str, Any]]] = {}
        for category, rows in services.items():
            name = str(category or "").strip()
            if not name:
                continue
            if not isinstance(rows, list):
                continue
            cleaned: List[Dict[str, Any]] = []
            for row in rows:
                if not isinstance(row, dict):
                    continue
                row_name = str(row.get("name", "")).strip()
                if not row_name:
                    continue
                cleaned.append(
                    {
                        "name": row_name,
                        "multiplier": float(row.get("multiplier", 1.0) or 0.0),
                        "is_base": bool(row.get("is_base", False)),
                    }
                )
            if cleaned:
                result[name] = cleaned
        if not result:
            return deepcopy(default_services)
        return result

    @classmethod
    def _load_config(cls) -> Tuple[List[Dict[str, Any]], Dict[str, List[Dict[str, Any]]]]:
        raw = cls._load_raw_config()
        defaults = cls._default_config()
        rows = cls._validate_rows(raw.get("translation_rows"), defaults["translation_rows"])
        additional = cls._validate_additional(
            raw.get("additional_services"), defaults["additional_services"]
        )
        return rows, additional

    @classmethod
    def reload(cls) -> None:
        rows, additional = cls._load_config()
        cls.TRANSLATION_ROWS = rows
        cls.ROW_NAMES = [row["name"] for row in rows]
        cls.ADDITIONAL_SERVICES = additional

    @classmethod
    def get_config(cls) -> Dict[str, Any]:
        return {
            "translation_rows": deepcopy(cls.TRANSLATION_ROWS),
            "additional_services": deepcopy(cls.ADDITIONAL_SERVICES),
        }

    @classmethod
    def save_config(
        cls,
        translation_rows: List[Dict[str, Any]],
        additional_services: Dict[str, List[Dict[str, Any]]],
    ) -> None:
        data = {
            "translation_rows": translation_rows,
            "additional_services": additional_services,
        }
        path = cls._config_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False, indent=2)
        cls.reload()

    @classmethod
    def reset_to_defaults(cls) -> None:
        path = cls._config_path()
        try:
            if path.exists():
                path.unlink()
        except OSError:
            # If deletion fails we still fall back to writing defaults below
            pass
        defaults = cls._default_config()
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", encoding="utf-8") as fh:
            json.dump(defaults, fh, ensure_ascii=False, indent=2)
        cls.reload()


# Ensure the in-memory configuration is initialized at import time.
ServiceConfig.reload()
