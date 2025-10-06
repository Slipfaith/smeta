"""Utilities for converting and formatting numeric values consistently.

This module centralises all operations for parsing and formatting numeric
values that are shared between the GUI layer and service layer.  The goal is
for widgets, exporters and other services to rely on the same canonical
implementation instead of keeping local copies of helpers.
"""
from __future__ import annotations

from typing import Union

NumberLike = Union[int, float, str]

_DEFAULT_RATE_DECIMALS = 3
_DEFAULT_AMOUNT_DECIMALS = 2


def rate_decimal_places() -> int:
    """Return the standard amount of decimal places for rate values."""
    return _DEFAULT_RATE_DECIMALS


def amount_decimal_places() -> int:
    """Return the standard amount of decimal places for monetary amounts."""
    return _DEFAULT_AMOUNT_DECIMALS


def decimal_separator_for_lang(lang: str) -> str:
    """Return the decimal separator preferred for *lang*.

    English interfaces default to a dot while all other languages use comma.
    """
    return "." if (lang or "").lower().startswith("en") else ","


def parse_number(value: NumberLike) -> float:
    """Convert *value* to ``float`` in a tolerant manner."""
    if isinstance(value, (int, float)):
        return float(value)
    text = (value or "0").strip()
    # allow grouping spaces before converting
    text = text.replace(" ", "")
    try:
        return float(text.replace(",", "."))
    except ValueError:
        return 0.0


def _resolve_separator(value: NumberLike, preferred: str | None) -> str:
    if preferred:
        return preferred
    if isinstance(value, str) and "," in value and "." not in value:
        return ","
    return "."


def format_rate(value: NumberLike, separator: str | None = None) -> str:
    """Format *value* as a rate string with canonical precision."""
    sep = _resolve_separator(value, separator)
    number = parse_number(value)
    text = f"{number:.{rate_decimal_places()}f}"
    if text.endswith("000"):
        text = text[:-1]
    elif text.endswith("00"):
        text = text[:-1]
    if sep == ",":
        text = text.replace(".", ",")
    return text


def format_amount(value: NumberLike, lang: str) -> str:
    """Format a monetary amount using locale specific separators."""
    number = parse_number(value)
    formatted = f"{number:,.{amount_decimal_places()}f}"
    if (lang or "").lower().startswith("en"):
        return formatted
    return formatted.replace(",", " ").replace(".", ",")


def convert_rate_value(value: NumberLike, multiplier: float, separator: str | None = None) -> str:
    """Multiply *value* by *multiplier* and return formatted rate text."""
    return format_rate(parse_number(value) * multiplier, separator)


def currency_suffix(symbol: str | None) -> str:
    return f" ({symbol})" if symbol else ""


def with_currency_suffix(text: str, symbol: str | None) -> str:
    return f"{text}{currency_suffix(symbol)}"


