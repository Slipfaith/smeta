"""Helpers for working with language and territory codes.

This module centralises the logic required to convert verbose language and
country names into standardised abbreviations.  The helpers are shared between
the GUI and the XML parser so that both parts of the application display
languages consistently (e.g. ``Malay (MY)`` instead of ``Malay (Malaysia)``).
"""

from __future__ import annotations

import csv
import re
from contextlib import suppress
from functools import lru_cache
from pathlib import Path
from typing import Dict

from babel import Locale
import pycountry


__all__ = [
    "RU_TERRITORY_ABBREVIATIONS",
    "country_to_code",
    "determine_short_code",
    "localise_territory_code",
    "replace_territory_with_code",
]


# Common short forms for Russian territory names that are not provided by Babel
# by default.  The keys are ISO alpha-2 codes and the values are the display
# strings that should be used in Russian UI contexts.
RU_TERRITORY_ABBREVIATIONS = {"US": "США", "RU": "РФ"}


_CODE_PATTERN = re.compile(r"^[A-Za-z]{2,3}$|^\d{3}$")
_PAREN_RE = re.compile(r"\(([^()]+)\)")


def _normalize(text: str) -> str:
    """Normalize strings for dictionary lookups."""

    return re.sub(r"\s+", " ", text.strip()).lower()


@lru_cache(maxsize=None)
def _alpha3_to_alpha2_map() -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    for country in pycountry.countries:
        alpha2 = getattr(country, "alpha_2", "")
        alpha3 = getattr(country, "alpha_3", "")
        if alpha2 and alpha3:
            mapping[alpha3.upper()] = alpha2.upper()
    return mapping


def _territory_from_code(code: str) -> str:
    """Extract a territory code from a locale identifier.

    Only ISO alpha-2 territory codes (and numeric UN M.49 codes) are
    considered.  Script subtags such as ``Latn`` are ignored because they do
    not represent countries and therefore should not be displayed in place of
    a territory abbreviation.
    """

    if not code:
        return ""

    normalized = code.replace("_", "-").strip()
    if not normalized:
        return ""

    parts = [part for part in normalized.split("-") if part]
    if len(parts) <= 1:
        return ""

    for part in reversed(parts[1:]):
        part_upper = part.upper()
        if len(part_upper) == 2 and part_upper.isalpha():
            return part_upper
        if part.isdigit():
            return part
        if len(part_upper) == 3 and part_upper.isalpha():
            alpha2 = _alpha3_to_alpha2_map().get(part_upper)
            if alpha2:
                return alpha2
    return ""


def _pycountry_country_lookup(name: str) -> str:
    if not name:
        return ""
    try:
        country = pycountry.countries.lookup(name)
    except LookupError:
        return ""
    code = getattr(country, "alpha_2", "") or getattr(country, "alpha_3", "")
    return code.upper() if code else ""


def _pycountry_language_lookup(value: str) -> str:
    if not value:
        return ""
    try:
        lang = pycountry.languages.lookup(value)
    except LookupError:
        return ""
    for attr in ("alpha_2", "alpha_3", "bibliographic", "terminology"):
        code = getattr(lang, attr, "")
        if code:
            return code.upper()
    return ""


@lru_cache(maxsize=None)
def _country_alias_map() -> Dict[str, str]:
    """Build a mapping of country names (RU/EN) to ISO codes.

    The map is seeded from ``languages/languages.csv`` so that any custom
    localisations present in the project are respected.
    """

    mapping: Dict[str, str] = {}
    csv_path = Path(__file__).resolve().parents[1] / "languages" / "languages.csv"

    try:
        with open(csv_path, encoding="utf-8-sig") as f:
            reader = csv.DictReader(f, delimiter=";")
            for row in reader:
                code = row.get("Код", "").strip()
                country_en = row.get("Страна (EN)", "").strip()
                country_ru = row.get("Страна (RU)", "").strip()

                territory = _territory_from_code(code)
                if not territory:
                    territory = _pycountry_country_lookup(country_en)
                if not territory:
                    territory = _pycountry_country_lookup(country_ru)

                if not territory:
                    continue

                if country_en:
                    mapping[_normalize(country_en)] = territory
                if country_ru:
                    mapping[_normalize(country_ru)] = territory

                # Names like "Latin, Azerbaijan" should also map to the
                # underlying territory.  Adding the suffixes ensures lookups
                # for either component succeed.
                if country_en and "," in country_en:
                    parts = [p.strip() for p in country_en.split(",") if p.strip()]
                    for part in parts[1:]:
                        mapping.setdefault(_normalize(part), territory)
                if country_ru and "," in country_ru:
                    parts = [p.strip() for p in country_ru.split(",") if p.strip()]
                    for part in parts[1:]:
                        mapping.setdefault(_normalize(part), territory)
    except FileNotFoundError:
        pass

    return mapping


@lru_cache(maxsize=None)
def _babel_country_map() -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    for lang in ("en", "ru"):
        try:
            locale = Locale(lang)
        except Exception:
            continue
        with suppress(Exception):
            for code, display in locale.territories.items():
                mapping[_normalize(display)] = code.upper()
        with suppress(Exception):
            short_map = locale._data.get("short_territories", {})
            for code, display in short_map.items():
                mapping[_normalize(display)] = code.upper()

    for code, short in RU_TERRITORY_ABBREVIATIONS.items():
        mapping[_normalize(short)] = code

    # Common unofficial abbreviations.
    mapping.setdefault("uk", "GB")

    return mapping


def country_to_code(name: str) -> str:
    """Return ISO code for a country or territory name.

    Works with English and Russian names as well as existing abbreviations.
    Returns an empty string when the input cannot be resolved.
    """

    if not name:
        return ""

    stripped = name.strip()
    if _CODE_PATTERN.fullmatch(stripped):
        return stripped.upper()

    norm = _normalize(stripped)

    mapping = _country_alias_map()
    if norm in mapping:
        return mapping[norm]

    babel_map = _babel_country_map()
    if norm in babel_map:
        return babel_map[norm]

    code = _pycountry_country_lookup(stripped)
    if code:
        return code

    return ""


def _language_code_from_row(code: str, lang_en: str) -> str:
    candidates = []
    if code:
        candidates.append(code.strip())
        base = code.split("-")[0].strip()
        if base:
            candidates.append(base)
    if lang_en:
        candidates.append(lang_en)

    seen = set()
    for candidate in candidates:
        candidate_norm = candidate.strip()
        if not candidate_norm or candidate_norm.lower() in seen:
            continue
        seen.add(candidate_norm.lower())

        code_value = _pycountry_language_lookup(candidate_norm)
        if code_value:
            return code_value

    if candidates:
        base = candidates[0].split("-")[0]
        if len(base) in (2, 3) and base.isalpha():
            return base.upper()

    return ""


def determine_short_code(
    code: str, lang_en: str, country_en: str, country_ru: str
) -> str:
    """Determine an appropriate short code for a CSV language row."""

    territory = _territory_from_code(code)
    if territory:
        return territory

    for value in (country_en, country_ru):
        territory = country_to_code(value)
        if territory:
            return territory

    if code and "-" in code:
        # Locale contains a script subtag but no territory.  In this case we
        # avoid adding a language code suffix to keep the label concise.
        return ""

    return _language_code_from_row(code, lang_en)


def localise_territory_code(code: str, lang: str) -> str:
    """Return a locale-aware representation of a territory code."""

    if not code:
        return ""
    if lang.lower().startswith("ru"):
        return RU_TERRITORY_ABBREVIATIONS.get(code, code)
    return code


def replace_territory_with_code(text: str, lang: str) -> str:
    """Replace the territory part of ``text`` (inside parentheses) with a code."""

    if not text:
        return ""

    def _repl(match: re.Match[str]) -> str:
        territory_name = match.group(1).strip()
        code = country_to_code(territory_name)
        if not code:
            return match.group(0)
        return f"({localise_territory_code(code, lang)})"

    return _PAREN_RE.sub(_repl, text, count=1)

