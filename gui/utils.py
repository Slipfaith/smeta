import re
from typing import Union

from logic.language_codes import apply_territory_overrides


def format_rate(value: Union[int, float, str], sep: str | None = None) -> str:
    """Format rate according to GUI rules."""

    if isinstance(value, str):
        raw = value.strip()
        if sep is None:
            sep = "," if "," in raw and "." not in raw else "."
        num = float(raw.replace(",", "."))
    else:
        num = float(value)
        sep = sep or "."
    text = f"{num:.3f}"
    if text.endswith("000"):
        text = text[:-1]
    elif text.endswith("00"):
        text = text[:-1]
    if sep == ",":
        text = text.replace(".", ",")
    return text


def format_amount(value: float, lang: str) -> str:
    """Format monetary values with locale-aware separators."""

    if lang == "en":
        return f"{value:,.2f}"
    return f"{value:,.2f}".replace(",", " ").replace(".", ",")


def _to_float(value: str) -> float:
    """Safely convert string to float."""

    try:
        return float((value or "0").replace(",", "."))
    except ValueError:
        return 0.0


def format_language_display(text: str, locale: str) -> str:
    """Apply locale-specific tweaks to language names displayed in the UI."""

    if not text:
        return ""

    formatted = apply_territory_overrides(text, locale)
    return re.sub(r"\s{2,}", " ", formatted).strip()
