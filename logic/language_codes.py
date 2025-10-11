"""Helpers for working with language and territory codes.

This module centralises the logic required to convert verbose language and
country names into standardised abbreviations.  The helpers are shared between
the GUI and the XML parser so that both parts of the application display
languages consistently (e.g. ``Malay (MY)`` instead of ``Malay (Malaysia)``).
"""

from __future__ import annotations

import re
from contextlib import suppress
from functools import lru_cache
from typing import Dict, Tuple

from babel import Locale
import langcodes


__all__ = [
    "RU_TERRITORY_ABBREVIATIONS",
    "country_to_code",
    "determine_short_code",
    "apply_territory_overrides",
]


# Common short forms for Russian territory names that are not provided by Babel
# by default.  The keys are ISO alpha-2 codes and the values are the display
# strings that should be used in Russian UI contexts.
RU_TERRITORY_ABBREVIATIONS = {"US": "США", "RU": "РФ"}


@lru_cache(maxsize=None)
def _locale_territory_info(lang: str) -> Tuple[Dict[str, str], Dict[str, str]]:
    """Return full and short territory names for ``lang`` locale."""

    if not lang:
        return {}, {}

    normalized = lang.replace("-", "_")

    try:
        locale = Locale.parse(normalized)
    except Exception:
        base = normalized.split("_")[0]
        if not base:
            return {}, {}
        try:
            locale = Locale.parse(base)
        except Exception:
            return {}, {}

    territories: Dict[str, str] = {}
    short_map: Dict[str, str] = {}

    with suppress(Exception):
        territories = {
            code.upper(): name
            for code, name in locale.territories.items()
            if code and name
        }

    with suppress(Exception):
        short_data = locale._data.get("short_territories", {})
        short_map = {
            code.upper(): name
            for code, name in short_data.items()
            if code and name
        }

    return territories, short_map


_CODE_PATTERN = re.compile(r"^[A-Za-z]{2,3}$|^\d{3}$")
_PAREN_RE = re.compile(r"\(([^()]+)\)")
_TERRITORY_DISPLAY_OVERRIDES = {
    "en": {"Latin America": "Latam"},
    "ru": {"Латинская Америка": "Латам"},
}


def _normalise_lang(lang: str) -> str:
    return (lang or "").split("-")[0].split("_")[0].lower()


def apply_territory_overrides(text: str, lang: str) -> str:
    """Apply locale-specific replacements and filtering to territory names."""

    if not text:
        return ""

    stripped = text.strip()
    if not stripped:
        return ""

    language = _normalise_lang(lang)
    overrides = _TERRITORY_DISPLAY_OVERRIDES.get(language, {})

    def _repl(match: re.Match[str]) -> str:
        inner = match.group(1).strip()
        replacement = overrides.get(inner)
        if replacement:
            return f" ({replacement})"
        return match.group(0)

    replaced = _PAREN_RE.sub(_repl, stripped)
    return re.sub(r"\s{2,}", " ", replaced).strip()


def _normalize(text: str) -> str:
    """Normalize strings for dictionary lookups."""

    return re.sub(r"\s+", " ", text.strip()).lower()


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

    try:
        lang = langcodes.Language.get(normalized)
    except langcodes.LanguageTagError:
        parts = [part for part in normalized.split("-") if part]
        if len(parts) == 1 and len(parts[0]) == 2 and parts[0].isalpha():
            return parts[0].upper()
        return ""

    territory = lang.territory or ""
    if territory:
        return territory.upper()

    base_language = (lang.language or "").lower()
    if not base_language:
        parts = [part for part in normalized.split("-") if part]
        if parts:
            base_language = parts[0].lower()

    # Try to extract territory from maximised tag (e.g. ``pt`` -> ``PT``)
    maximized = lang.maximize()
    if maximized.territory:
        max_language = (maximized.language or "").lower()
        if base_language and max_language and max_language != base_language:
            if lang.script and base_language:
                with suppress(langcodes.LanguageTagError):
                    base_maximized = langcodes.Language.get(base_language).maximize()
                    base_territory = base_maximized.territory or ""
                    if (
                        base_territory
                        and (base_maximized.language or "").lower() == base_language
                    ):
                        return base_territory.upper()
            return ""
        return maximized.territory.upper()

    return ""


@lru_cache(maxsize=None)
def _territory_name_map() -> Dict[str, str]:
    mapping: Dict[str, str] = {}

    for lang in ("en", "ru"):
        try:
            locale = Locale(lang)
        except Exception:
            continue

        with suppress(Exception):
            for code, display in locale.territories.items():
                if not code or not display:
                    continue
                mapping[_normalize(display)] = code.upper()
                for part in re.split(r"[,/]| - ", display):
                    cleaned = part.strip()
                    if cleaned and cleaned != display:
                        mapping.setdefault(_normalize(cleaned), code.upper())

        with suppress(Exception):
            short_map = locale._data.get("short_territories", {})
            for code, display in short_map.items():
                if not code or not display:
                    continue
                mapping[_normalize(display)] = code.upper()

    for code, short in RU_TERRITORY_ABBREVIATIONS.items():
        mapping[_normalize(short)] = code

    # Include override display names (e.g. "Latam") so that abbreviated
    # variants resolve to the same territory codes as their full forms.
    for override_map in _TERRITORY_DISPLAY_OVERRIDES.values():
        for original, alias in override_map.items():
            alias_norm = _normalize(alias)
            if alias_norm in mapping:
                continue
            code = mapping.get(_normalize(original))
            if code:
                mapping[alias_norm] = code

    # Allow matching against plain codes ("US", "gb", "410").
    for code in list(mapping.values()):
        mapping.setdefault(_normalize(code), code)

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

    mapping = _territory_name_map()
    if norm in mapping:
        return mapping[norm]

    # Try to parse the string as a language tag and reuse the territory part.
    try:
        lang = langcodes.Language.get(stripped)
        if lang.territory:
            return lang.territory.upper()
    except langcodes.LanguageTagError:
        pass

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

        try:
            lang = langcodes.Language.get(candidate_norm)
            if lang.language:
                return lang.language.upper()
        except langcodes.LanguageTagError:
            try:
                match = langcodes.find(candidate_norm)
            except LookupError:
                continue
            try:
                lang = langcodes.Language.get(match)
                if lang.language:
                    return lang.language.upper()
            except langcodes.LanguageTagError:
                continue

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


