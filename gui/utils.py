from typing import Union
import re

from babel import Locale


RU_SHORT_TERRITORIES = {
    "US": "США",
    "RU": "РФ",
}


def format_rate(value: Union[int, float]) -> str:
    """Format rate according to GUI rules.

    Numbers are formatted with up to three decimals. If the formatted
    string ends with two or three zeros after the decimal separator,
    trim one trailing zero so that ``3.300`` becomes ``3.30`` and
    ``3.000`` becomes ``3.00``.
    """
    num = float(value)
    text = f"{num:.3f}"
    if text.endswith("000"):
        return text[:-1]
    if text.endswith("00"):
        return text[:-1]
    return text


def shorten_locale(text: str, lang: str) -> str:
    """Shorten territory name in parentheses using Babel data."""
    match = re.search(r"\(([^)]+)\)", text)
    if not match:
        return text
    country_full = match.group(1)
    try:
        locale = Locale(lang)
    except Exception:
        return text
    territories = locale.territories
    full_to_code = {name: code for code, name in territories.items()}
    code = full_to_code.get(country_full)
    if not code:
        return text
    short_map = locale._data.get("short_territories", {})
    country_short = short_map.get(code)
    if not country_short:
        if lang == "ru":
            country_short = RU_SHORT_TERRITORIES.get(code, country_full)
        else:
            country_short = code
    return text[: match.start() + 1] + country_short + text[match.end() - 1 :]
