import json
import os
import platform
from typing import List, Dict, Optional, Tuple

APP_NAME = "ProjectCalculator"
LANG_FILE = "languages.json"


def _appdata_base() -> str:
    system = platform.system().lower()
    home = os.path.expanduser("~")
    if "windows" in system:
        base = os.environ.get("APPDATA") or os.path.join(home, "AppData", "Roaming")
        return base
    if "darwin" in system:  # macOS
        return os.path.join(home, "Library", "Application Support")
    # linux/other
    return os.path.join(home, ".config")


def get_appdata_dir() -> str:
    base = _appdata_base()
    path = os.path.join(base, APP_NAME)
    os.makedirs(path, exist_ok=True)
    return path


def _default_languages() -> List[Dict[str, str]]:
    # Без кодов — только названия на RU и EN
    return [
        {"en": "English",               "ru": "Английский"},
        {"en": "Russian",               "ru": "Русский"},
        {"en": "Chinese (Simplified)",  "ru": "Китайский (Упрощенный)"},
        {"en": "Chinese (Traditional)", "ru": "Китайский (традиц.)"},
        {"en": "German",                "ru": "Немецкий"},
        {"en": "French",                "ru": "Французский"},
        {"en": "Spanish",               "ru": "Испанский"},
        {"en": "Portuguese",            "ru": "Португальский"},
        {"en": "Italian",               "ru": "Итальянский"},
        {"en": "Japanese",              "ru": "Японский"},
        {"en": "Korean",                "ru": "Корейский"},
        {"en": "Arabic",                "ru": "Арабский"},
        {"en": "Ukrainian",             "ru": "Украинский"},
        {"en": "Polish",                "ru": "Польский"},
        {"en": "Dutch",                 "ru": "Нидерландский"},
        {"en": "Turkish",               "ru": "Турецкий"},
        {"en": "Czech",                 "ru": "Чешский"},
        {"en": "Slovak",                "ru": "Словацкий"},
        {"en": "Romanian",              "ru": "Румынский"},
        {"en": "Bulgarian",             "ru": "Болгарский"},
        {"en": "Hungarian",             "ru": "Венгерский"},
        {"en": "Greek",                 "ru": "Греческий"},
        {"en": "Hebrew",                "ru": "Иврит"},
        {"en": "Hindi",                 "ru": "Хинди"},
        {"en": "Thai",                  "ru": "Тайский"},
        {"en": "Vietnamese",            "ru": "Вьетнамский"},
        {"en": "Indonesian",            "ru": "Индонезийский"},
        {"en": "Malay",                 "ru": "Малайский"},
        {"en": "Finnish",               "ru": "Финский"},
        {"en": "Swedish",               "ru": "Шведский"},
        {"en": "Norwegian",             "ru": "Норвежский"},
        {"en": "Danish",                "ru": "Датский"},
        {"en": "Estonian",              "ru": "Эстонский"},
        {"en": "Latvian",               "ru": "Латышский"},
        {"en": "Valyrian",              "ru": "Валирийский"},
        {"en": "Georgian",              "ru": "Грузинский"},
        {"en": "Armenian",              "ru": "Армянский"},
        {"en": "Azerbaijani",           "ru": "Азербайджанский"},
        {"en": "Kazakh",                "ru": "Казахский"},
        {"en": "Uzbek",                 "ru": "Узбекский"},
        {"en": "Belarusian",            "ru": "Белорусский"},
        {"en": "Serbian",               "ru": "Сербский"},
        {"en": "Croatian",              "ru": "Хорватский"},
        {"en": "Bosnian",               "ru": "Боснийский"},
        {"en": "Slovenian",             "ru": "Словенский"},
        {"en": "Macedonian",            "ru": "Македонский"},
        {"en": "Catalan",               "ru": "Каталанский"},
    ]


def _languages_path() -> str:
    return os.path.join(get_appdata_dir(), LANG_FILE)


def ensure_languages_file() -> str:
    """Гарантирует наличие languages.json с дефолтным списком."""
    path = _languages_path()
    if not os.path.exists(path):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(_default_languages(), f, ensure_ascii=False, indent=2)
    return path


def _norm_pair(en: str, ru: str) -> Tuple[str, str]:
    return (str(en or "").strip().lower(), str(ru or "").strip().lower())


def load_languages() -> List[Dict[str, str]]:
    ensure_languages_file()
    path = _languages_path()
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return _default_languages()

    if isinstance(data, list):
        out: List[Dict[str, str]] = []
        seen = set()
        for it in data:
            if not isinstance(it, dict):
                continue
            en = str(it.get("en", "")).strip()
            ru = str(it.get("ru", "")).strip()
            # хотя бы одно название должно быть
            if not (en or ru):
                continue
            # автозаполнение недостающего
            if not en:
                en = ru
            if not ru:
                ru = en
            key = _norm_pair(en, ru)
            if key not in seen:
                out.append({"en": en, "ru": ru})
                seen.add(key)
        return out

    return _default_languages()


def save_languages(langs: List[Dict[str, str]]) -> bool:
    path = _languages_path()
    try:
        # лёгкая нормализация перед сохранением
        cleaned: List[Dict[str, str]] = []
        seen = set()
        for it in langs:
            en = str(it.get("en", "")).strip()
            ru = str(it.get("ru", "")).strip()
            if not (en or ru):
                continue
            if not en:
                en = ru
            if not ru:
                ru = en
            key = _norm_pair(en, ru)
            if key in seen:
                continue
            cleaned.append({"en": en, "ru": ru})
            seen.add(key)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(cleaned, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def add_language(en: str, ru: str) -> bool:
    """
    Добавляет/обновляет язык в локальный конфиг (без кодов).
    Совпадение ищется по паре (en, ru) без учёта регистра; если один из них пуст —
    ищем по имеющемуся.
    """
    en = (en or "").strip()
    ru = (ru or "").strip()
    if not (en or ru):
        return False

    langs = load_languages()

    # Нормализуем вход
    if not en:
        en = ru
    if not ru:
        ru = en

    key_new = _norm_pair(en, ru)

    # ищем существующую запись — либо точным совпадением пары,
    # либо совпадением одного из названий (чтобы не плодить дубликаты)
    idx: Optional[int] = None
    for i, l in enumerate(langs):
        k = _norm_pair(l.get("en", ""), l.get("ru", ""))
        if k == key_new or en.lower() == l.get("en", "").strip().lower() or ru.lower() == l.get("ru", "").strip().lower():
            idx = i
            break

    new_entry = {"en": en, "ru": ru}
    if idx is None:
        langs.append(new_entry)
    else:
        langs[idx] = new_entry

    return save_languages(langs)
