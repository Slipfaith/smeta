import re
from babel import Locale

from logic.language_codes import RU_TERRITORY_ABBREVIATIONS


RU_SHORT_TERRITORIES = RU_TERRITORY_ABBREVIATIONS


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
