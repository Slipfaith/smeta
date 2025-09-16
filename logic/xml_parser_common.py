from __future__ import annotations

import csv
import re
from pathlib import Path
from typing import Dict

import langcodes
import pycountry

from .language_codes import (
    determine_short_code,
    localise_territory_code,
    replace_territory_with_code,
)


LANGUAGE_CODE_MAP: Dict[str, str] = {}
LANGUAGE_NAME_MAP: Dict[str, str] = {}


def _load_languages_csv() -> None:
    """Загружает сопоставления кодов языков из languages/languages.csv."""

    csv_path = Path(__file__).resolve().parents[1] / "languages" / "languages.csv"

    try:
        with open(csv_path, encoding="utf-8-sig") as f:
            reader = csv.DictReader(f, delimiter=";")
            for row in reader:
                code = row.get("Код", "").strip().lower()
                lang_en = row.get("Язык (EN)", "").strip()
                country_en = row.get("Страна (EN)", "").strip()
                lang_ru = row.get("Язык (RU)", "").strip()
                country_ru = row.get("Страна (RU)", "").strip()

                if not lang_en:
                    continue

                country_code = determine_short_code(code, lang_en, country_en, country_ru)
                display_en = f"{lang_en} ({country_code})" if country_code else lang_en

                lang_ru_final = lang_ru if re.search("[А-Яа-я]", lang_ru) else ""
                country_ru_final = (
                    country_ru if re.search("[А-Яа-я]", country_ru) else ""
                )

                if not lang_ru_final:
                    try:
                        lang_code = code or langcodes.find(lang_en)
                        lang_ru_final = langcodes.Language.get(lang_code).language_name("ru")
                    except Exception:
                        lang_ru_final = lang_en

                if not country_ru_final and code and "-" in code:
                    try:
                        country_ru_final = langcodes.Language.get(code).territory_name("ru")
                    except Exception:
                        country_ru_final = country_en

                country_code_ru = localise_territory_code(country_code, "ru")
                if country_code_ru:
                    display_ru = f"{lang_ru_final} ({country_code_ru})"
                elif country_ru_final:
                    display_ru = f"{lang_ru_final} ({country_ru_final})"
                else:
                    display_ru = lang_ru_final

                display = display_ru or display_en
                if display:
                    display = display[0].upper() + display[1:]

                if code:
                    LANGUAGE_CODE_MAP[code] = display

                for key in filter(
                    None,
                    [
                        lang_en.lower(),
                        display_en.lower(),
                        f"{lang_en.lower()} ({country_en.lower()})"
                        if lang_en and country_en
                        else None,
                        lang_ru.lower(),
                        display_ru.lower(),
                        f"{lang_ru.lower()} ({country_ru.lower()})"
                        if lang_ru and country_ru
                        else None,
                        f"{lang_ru.lower()} ({country_ru_final.lower()})"
                        if lang_ru and country_ru_final
                        else None,
                    ],
                ):
                    LANGUAGE_NAME_MAP[key] = display
    except FileNotFoundError:
        # Файл со списком языков отсутствует — будем использовать только стандартные методы
        pass


_load_languages_csv()


def lookup_language(value: str) -> str:
    """Возвращает отображаемое название языка из CSV по коду или имени."""

    if not value:
        return ""

    norm = value.strip().lower()
    code_key = norm.replace("_", "-")
    return LANGUAGE_CODE_MAP.get(code_key) or LANGUAGE_NAME_MAP.get(norm, "")


def norm_lang(code: str) -> str:
    if not code:
        return ""
    return code.split("-")[0].upper()


def _display_with_pycountry(normalized: str) -> str:
    language_code = normalized
    territory_part = ""
    if "-" in normalized:
        language_code, territory_part = normalized.split("-", 1)

    try:
        language = pycountry.languages.lookup(language_code)
    except LookupError:
        return ""

    base_code = getattr(language, "alpha_2", "") or getattr(language, "alpha_3", "")

    if base_code:
        try:
            code_to_use = base_code
            if territory_part:
                code_to_use = f"{base_code}-{territory_part.upper()}"
            result = langcodes.Language.get(code_to_use).display_name("ru")
            return replace_territory_with_code(result, "ru")
        except langcodes.LanguageTagError:
            pass

    name = getattr(language, "name", "") or getattr(language, "common_name", "")
    if not name:
        names = getattr(language, "names", None)
        if names:
            name = names[0]

    territory_display = ""
    if territory_part:
        territory_display = localise_territory_code(territory_part, "ru")
        if not territory_display:
            try:
                territory = pycountry.countries.lookup(territory_part)
                territory_display = territory.name
            except LookupError:
                territory_display = territory_part.upper()

    if name and territory_display:
        return f"{name} ({territory_display})"
    return name or ""


def expand_language_code(code: str) -> str:
    """Преобразует языковой код в человекочитаемое название (на русском)."""

    if not code:
        return ""

    normalized = code.replace("_", "-")

    csv_name = lookup_language(normalized)
    if csv_name:
        return csv_name

    pycountry_display = _display_with_pycountry(normalized)
    if pycountry_display:
        return pycountry_display

    try:
        result = langcodes.Language.get(normalized).display_name("ru")
        return replace_territory_with_code(result, "ru")
    except langcodes.LanguageTagError:
        simple_code = norm_lang(normalized)
        return simple_code


def normalize_language_name(name: str) -> str:
    """Нормализует название или код языка и возвращает его на русском."""

    if not name:
        return ""

    name = name.strip()

    csv_name = lookup_language(name)
    if csv_name:
        return csv_name

    try:
        result = langcodes.Language.get(name).display_name("ru")
        return replace_territory_with_code(result, "ru")
    except langcodes.LanguageTagError:
        pass

    try:
        if "(" in name and ")" in name:
            lang_part, region_part = name.split("(", 1)
            lang_part = lang_part.strip()
            region_part = region_part.strip(") ").strip()

            if region_part.lower() in {"simplified", "traditional"}:
                code = "zh-Hans" if region_part.lower() == "simplified" else "zh-Hant"
            else:
                lang = pycountry.languages.lookup(lang_part)
                country = pycountry.countries.lookup(region_part)
                code = f"{lang.alpha_2}-{country.alpha_2}"
        else:
            lang = pycountry.languages.lookup(name)
            code = getattr(lang, "alpha_2", "") or getattr(lang, "alpha_3", "")

        if code:
            result = langcodes.Language.get(code).display_name("ru")
            return replace_territory_with_code(result, "ru")
    except LookupError:
        try:
            code = langcodes.find(name)
            result = langcodes.Language.get(code).display_name("ru")
            return replace_territory_with_code(result, "ru")
        except Exception:
            return ""
    except langcodes.LanguageTagError:
        return ""

    return ""


def resolve_language_display(value: str) -> str:
    value = value.strip()
    if not value:
        return ""

    display = expand_language_code(value)
    if display:
        return display

    normalized = normalize_language_name(value)
    if normalized:
        return normalized

    if re.fullmatch(r"[A-Za-z]{2,3}(?:-[A-Za-z]{2,3})?", value):
        return value.upper()

    return value
