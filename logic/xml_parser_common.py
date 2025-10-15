from __future__ import annotations

import re
from typing import Iterable, Set, Tuple

import langcodes

from .language_codes import apply_territory_overrides, country_to_code


_SCRIPT_HINTS = {
    "simplified": "Hans",
    "traditional": "Hant",
    "упрощ": "Hans",
    "традиц": "Hant",
}


def norm_lang(code: str) -> str:
    if not code:
        return ""
    return code.replace("_", "-").split("-")[0].upper()


def _format_display(text: str, locale: str) -> str:
    if not text:
        return ""
    stripped = text.strip()
    if not stripped:
        return ""
    formatted = stripped[0].upper() + stripped[1:]
    return apply_territory_overrides(formatted, locale)


def _resolve_script_hint(hint: str) -> str:
    lowered = hint.lower().strip().rstrip(".")
    for key, script in _SCRIPT_HINTS.items():
        if lowered == key or lowered.startswith(f"{key} ") or lowered.startswith(key):
            return script
    return ""


def _language_tag_from_parts(base: str, region: str) -> str:
    try:
        base_tag = langcodes.find(base)
    except LookupError:
        return ""

    base_lang = langcodes.Language.get(base_tag)
    kwargs = {"language": base_lang.language or base_tag}
    if base_lang.script:
        kwargs["script"] = base_lang.script
    if base_lang.territory:
        kwargs["territory"] = base_lang.territory

    script_hint = _resolve_script_hint(region)
    if script_hint:
        try:
            kwargs["script"] = script_hint
            tag = langcodes.Language.make(**kwargs).to_tag()
        except langcodes.LanguageTagError:
            return ""
        try:
            if langcodes.Language.get(tag).is_valid():
                return tag
        except langcodes.LanguageTagError:
            return ""

    territory_code = country_to_code(region)
    if territory_code:
        try:
            kwargs["territory"] = territory_code
            tag = langcodes.Language.make(**kwargs).to_tag()
        except langcodes.LanguageTagError:
            return ""
        try:
            if langcodes.Language.get(tag).is_valid():
                return tag
        except langcodes.LanguageTagError:
            return ""

    return ""


def _language_tag_from_value(value: str) -> str:
    if not value:
        return ""

    stripped = value.strip()
    if not stripped:
        return ""

    normalized = stripped.replace("_", "-")

    try:
        candidate = langcodes.standardize_tag(normalized)
    except Exception:
        candidate = ""
    if candidate:
        try:
            if langcodes.Language.get(candidate).is_valid():
                return candidate
        except langcodes.LanguageTagError:
            candidate = ""

    match = re.match(r"(.+?)\s*\(([^()]+)\)$", stripped)
    if match:
        tag = _language_tag_from_parts(match.group(1).strip(), match.group(2).strip())
        if tag:
            return tag

    match = re.search(
        r"(?<![A-Za-z])([A-Za-z]{2,3}(?:-[A-Za-z0-9]{2,3})?)(?![A-Za-z])",
        stripped,
    )
    if match and len(stripped) <= 10:
        try:
            return langcodes.standardize_tag(match.group(1))
        except Exception:
            pass

    for sep in (",", "/", "-", "→"):
        if sep in stripped:
            parts = [part.strip() for part in stripped.split(sep) if part.strip()]
            if len(parts) == 2:
                tag = _language_tag_from_parts(parts[0], parts[1])
                if tag:
                    return tag

    try:
        return langcodes.find(stripped)
    except LookupError:
        return ""


def expand_language_code(code: str, locale: str = "ru") -> str:
    """Преобразует языковой код в человекочитаемое название."""

    if not code:
        return ""

    normalized = code.replace("_", "-")

    try:
        tag = _language_tag_from_value(normalized) or normalized
        language = langcodes.Language.get(tag)
        if not language.is_valid():
            raise langcodes.LanguageTagError(f"Invalid tag: {tag}")
        result = language.display_name(locale)
        return _format_display(result, locale)
    except langcodes.LanguageTagError:
        return norm_lang(normalized)


def language_identity(value: str) -> Tuple[str, str, str]:
    """Return normalized language, script and territory codes for *value*.

    The helper attempts to interpret ``value`` as a language identifier using
    the same heuristics as :func:`expand_language_code`.  The return value is a
    tuple ``(language, script, territory)`` where each element is normalised to
    lower/upper case respectively.  Empty strings indicate that the component
    could not be determined.
    """

    if not value:
        return "", "", ""

    tag = _language_tag_from_value(value)
    candidates = [candidate for candidate in (tag, value) if candidate]

    for candidate in candidates:
        try:
            lang = langcodes.Language.get(candidate)
        except langcodes.LanguageTagError:
            continue

        if not lang.is_valid():
            continue

        language = (lang.language or "").lower()
        script = (lang.script or "").title() if lang.script else ""
        territory = (lang.territory or "").upper()
        return language, script, territory

    return "", "", ""


def expand_target_matches(
    matched_norms: Iterable[str],
    available_norms: Iterable[str],
    source_norm: str = "",
) -> Set[str]:
    """Expand *matched_norms* with other variants of the same base language.

    The helper ensures that all target languages that belong to the same base
    language as items from *matched_norms* are selected.  The source language
    itself is removed from the result when ``source_norm`` is provided.
    """

    expanded_norms: Set[str] = set()
    base_languages: Set[str] = set()

    for norm in matched_norms:
        if not norm:
            continue
        expanded_norms.add(norm)
        base_languages.add(norm.split("-", 1)[0])

    if base_languages:
        for norm in available_norms:
            if not norm:
                continue
            base = norm.split("-", 1)[0]
            if base in base_languages:
                expanded_norms.add(norm)

    if source_norm:
        expanded_norms.discard(source_norm)

    return expanded_norms


def normalize_language_name(name: str, locale: str = "ru") -> str:
    """Нормализует название или код языка и возвращает его перевод."""

    if not name:
        return ""

    name = name.strip()

    try:
        tag = _language_tag_from_value(name)
        if not tag:
            return ""
        result = langcodes.Language.get(tag).display_name(locale)
        return _format_display(result, locale)
    except langcodes.LanguageTagError:
        try:
            match = langcodes.find(name)
        except LookupError:
            return ""
        try:
            result = langcodes.Language.get(match).display_name(locale)
            return _format_display(result, locale)
        except langcodes.LanguageTagError:
            return ""


def resolve_language_display(value: str, locale: str = "ru") -> str:
    value = value.strip()
    if not value:
        return ""

    display = expand_language_code(value, locale=locale)
    if display:
        return display

    normalized = normalize_language_name(value, locale=locale)
    if normalized:
        return normalized

    if re.fullmatch(r"[A-Za-z]{2,3}(?:-[A-Za-z]{2,3})?", value):
        return value.upper()

    return value
