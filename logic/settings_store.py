"""Persistent storage helpers for user-adjustable service settings."""

from __future__ import annotations

import json
from copy import deepcopy
from pathlib import Path
from typing import Any, Dict, Iterable, List

from .user_config import get_appdata_dir


SETTINGS_FILENAME = "settings.json"


def _settings_path() -> Path:
    """Return absolute path to the user settings JSON file."""

    path = Path(get_appdata_dir()) / SETTINGS_FILENAME
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


DEFAULT_TRANSLATION_ROWS: List[Dict[str, Any]] = [
    {"key": "new", "name": "Перевод, новые слова (100%)", "multiplier": 1.0, "is_base": True},
    {"key": "fuzzy_75_94", "name": "Перевод, совпадения 75-94% (66%)", "multiplier": 0.66, "is_base": False},
    {"key": "fuzzy_95_99", "name": "Перевод, совпадения 95-99% (33%)", "multiplier": 0.33, "is_base": False},
    {"key": "reps_100_30", "name": "Перевод, повторы и 100% совпадения (30%)", "multiplier": 0.30, "is_base": False},
]


DEFAULT_ADDITIONAL_SERVICES: Dict[str, List[Dict[str, Any]]] = {
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


DEFAULT_FUZZY_THRESHOLDS = {"new_words": 100, "others": 75}


DEFAULT_SETTINGS: Dict[str, Any] = {
    "translation_rows": DEFAULT_TRANSLATION_ROWS,
    "additional_services": DEFAULT_ADDITIONAL_SERVICES,
    "fuzzy_thresholds": DEFAULT_FUZZY_THRESHOLDS,
}


def _load_raw() -> Dict[str, Any]:
    path = _settings_path()
    if not path.exists():
        return {}
    try:
        with path.open("r", encoding="utf-8") as fh:
            data = json.load(fh)
        if isinstance(data, dict):
            return data
    except Exception:
        pass
    return {}


def _validate_multiplier(value: Any) -> float:
    try:
        num = float(value)
    except (TypeError, ValueError):
        return 0.0
    if num < 0:
        return 0.0
    if num > 100:
        return 100.0
    return num


def _normalize_translation_row(row: Dict[str, Any]) -> Dict[str, Any]:
    cleaned: Dict[str, Any] = {
        "key": row.get("key"),
        "name": str(row.get("name", "")).strip() or "Строка",
        "is_base": bool(row.get("is_base", False)),
        "multiplier": _validate_multiplier(row.get("multiplier", 0.0)),
    }
    return cleaned


def _normalize_translation_rows(rows: Iterable[Dict[str, Any]]) -> List[Dict[str, Any]]:
    cleaned: List[Dict[str, Any]] = []
    base_assigned = False
    for row in rows:
        if not isinstance(row, dict):
            continue
        norm = _normalize_translation_row(row)
        if norm["is_base"] and not base_assigned:
            base_assigned = True
        elif norm["is_base"] and base_assigned:
            norm["is_base"] = False
        cleaned.append(norm)
    if not cleaned:
        cleaned = [deepcopy(DEFAULT_TRANSLATION_ROWS[0])]
    if not any(row.get("is_base") for row in cleaned):
        cleaned[0]["is_base"] = True
    return cleaned


def _normalize_additional_services(data: Any) -> Dict[str, List[Dict[str, Any]]]:
    if not isinstance(data, dict):
        return deepcopy(DEFAULT_ADDITIONAL_SERVICES)
    cleaned: Dict[str, List[Dict[str, Any]]] = {}
    for section, rows in data.items():
        if not isinstance(section, str):
            continue
        section_name = section.strip() or "Категория"
        cleaned_rows: List[Dict[str, Any]] = []
        if isinstance(rows, list):
            for row in rows:
                if not isinstance(row, dict):
                    continue
                cleaned_rows.append(
                    {
                        "name": str(row.get("name", "")).strip() or "Услуга",
                        "multiplier": _validate_multiplier(row.get("multiplier", 0.0)),
                        "is_base": bool(row.get("is_base", False)),
                    }
                )
        if cleaned_rows:
            cleaned[section_name] = cleaned_rows
    return cleaned or deepcopy(DEFAULT_ADDITIONAL_SERVICES)


def _normalize_fuzzy_thresholds(data: Any) -> Dict[str, int]:
    defaults = DEFAULT_FUZZY_THRESHOLDS
    if not isinstance(data, dict):
        return dict(defaults)
    result: Dict[str, int] = {}
    for key in ("new_words", "others"):
        value = data.get(key, defaults.get(key, 0))
        try:
            number = int(value)
        except (TypeError, ValueError):
            number = defaults.get(key, 0)
        number = max(0, min(100, number))
        result[key] = number
    return result


def load_settings() -> Dict[str, Any]:
    """Return sanitized dictionary with all user settings."""

    raw = _load_raw()
    translation_rows = _normalize_translation_rows(raw.get("translation_rows", []))
    additional_services = _normalize_additional_services(raw.get("additional_services"))
    fuzzy_thresholds = _normalize_fuzzy_thresholds(raw.get("fuzzy_thresholds"))
    return {
        "translation_rows": translation_rows,
        "additional_services": additional_services,
        "fuzzy_thresholds": fuzzy_thresholds,
    }


def save_settings(settings: Dict[str, Any]) -> None:
    """Persist *settings* to disk after validation."""

    validated = {
        "translation_rows": _normalize_translation_rows(settings.get("translation_rows", [])),
        "additional_services": _normalize_additional_services(settings.get("additional_services")),
        "fuzzy_thresholds": _normalize_fuzzy_thresholds(settings.get("fuzzy_thresholds")),
    }
    path = _settings_path()
    with path.open("w", encoding="utf-8") as fh:
        json.dump(validated, fh, ensure_ascii=False, indent=2)


def get_translation_rows() -> List[Dict[str, Any]]:
    """Return a deep copy of translation rows configuration."""

    return deepcopy(load_settings()["translation_rows"])


def get_additional_services() -> Dict[str, List[Dict[str, Any]]]:
    """Return a deep copy of additional services configuration."""

    return deepcopy(load_settings()["additional_services"])


def get_fuzzy_thresholds() -> Dict[str, int]:
    """Return fuzzy threshold configuration."""

    return dict(load_settings()["fuzzy_thresholds"])


def update_translation_rows(rows: Iterable[Dict[str, Any]]) -> None:
    settings = load_settings()
    settings["translation_rows"] = list(rows)
    save_settings(settings)


def update_additional_services(data: Dict[str, List[Dict[str, Any]]]) -> None:
    settings = load_settings()
    settings["additional_services"] = data
    save_settings(settings)


def update_fuzzy_thresholds(new_words: int, others: int) -> None:
    settings = load_settings()
    settings["fuzzy_thresholds"] = {"new_words": new_words, "others": others}
    save_settings(settings)
