import unittest
from unittest.mock import patch

from logic import language_codes


class _DummyLanguage:
    def __init__(self, language="", territory="", script="", maximized=None):
        self.language = language
        self.territory = territory
        self.script = script
        self._maximized = maximized or self

    def maximize(self):
        return self._maximized


class TerritoryDetectionTests(unittest.TestCase):
    def test_script_only_code_ignores_mismatched_language_territory(self):
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

        self.assertEqual(short, "AZ")


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
