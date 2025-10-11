from __future__ import annotations

from unittest.mock import patch

from logic import language_codes
from logic.xml_parser_common import expand_language_code


class _DummyLanguage:
    def __init__(self, language="", territory="", script="", maximized=None):
        self.language = language
        self.territory = territory
        self.script = script
        self._maximized = maximized or self

    def maximize(self):
        return self._maximized


def test_country_to_code_returns_existing_code():
    assert language_codes.country_to_code("us") == "US"


def test_country_to_code_resolves_russian_abbreviation():
    assert language_codes.country_to_code("США") == "US"


def test_determine_short_code_prefers_territory_from_code():
    assert language_codes.determine_short_code("en-US", "", "", "") == "US"


def test_determine_short_code_falls_back_to_country_names():
    code = language_codes.determine_short_code("", "", "Germany", "")
    assert code == "DE"


def test_determine_short_code_uses_language_when_no_territory():
    code = language_codes.determine_short_code("", "Esperanto", "", "")
    assert code == "EO"


def test_determine_short_code_script_only_code_ignores_mismatched_language_territory():
    base_maximized = _DummyLanguage(language="uz", territory="UZ", script="Latn")
    base_language = _DummyLanguage(language="az", script="Latn", maximized=base_maximized)

    fallback_maximized = _DummyLanguage(language="az", territory="AZ")
    fallback_language = _DummyLanguage(language="az", maximized=fallback_maximized)

    def fake_get(tag):
        if tag == "az-Latn":
            return base_language
        if tag == "az":
            return fallback_language
        raise language_codes.langcodes.LanguageTagError

    with patch.object(language_codes.langcodes.Language, "get", side_effect=fake_get):
        short = language_codes.determine_short_code("az-Latn", "", "", "")

    assert short == "AZ"


def test_determine_short_code_returns_empty_for_script_only_when_no_fallback():
    def fake_get(tag):
        raise language_codes.langcodes.LanguageTagError

    with patch.object(language_codes.langcodes.Language, "get", side_effect=fake_get):
        short = language_codes.determine_short_code("zh-Hans", "", "", "")

    assert short == ""


def test_apply_territory_overrides_for_latin_america_en():
    assert (
        language_codes.apply_territory_overrides("Spanish (Latin America)", "en")
        == "Spanish (Latam)"
    )


def test_apply_territory_overrides_for_latin_america_ru():
    assert (
        language_codes.apply_territory_overrides(
            "испанский (Латинская Америка)", "ru"
        )
        == "испанский (Латам)"
    )


def test_apply_territory_overrides_ignores_unknown_locale():
    assert (
        language_codes.apply_territory_overrides("Spanish (Latin America)", "de")
        == "Spanish (Latin America)"
    )


def test_apply_territory_overrides_preserves_other_territories():
    assert (
        language_codes.apply_territory_overrides("English (United States)", "ru")
        == "English (United States)"
    )


def test_expand_language_code_preserves_territory_information():
    assert expand_language_code("zh-CN", locale="en").endswith("(China)")
    assert expand_language_code("zh-TW", locale="en").endswith("(Taiwan)")
