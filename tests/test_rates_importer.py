"""Tests for Excel rate matching helpers."""

from logic import rates_importer


def test_match_pairs_does_not_suggest_unavailable_targets() -> None:
    rates = {
        ("en", "zh-hans"): {"basic": 1.0, "complex": 2.0, "hour": 3.0},
        ("zh", "en"): {"basic": 1.5, "complex": 2.5, "hour": 3.5},
    }

    matches = rates_importer.match_pairs([("English", "Chinese")], rates)

    assert matches[0].excel_source == "English"
    assert matches[0].excel_target == ""
    assert matches[0].rates is None


def test_match_pairs_returns_exact_language_names_for_hits() -> None:
    rates = {
        ("en", "zh-hans"): {"basic": 1.0, "complex": 2.0, "hour": 3.0},
        ("zh", "en"): {"basic": 1.5, "complex": 2.5, "hour": 3.5},
    }

    matches = rates_importer.match_pairs(
        [("English", "Chinese (Simplified)")],
        rates,
    )

    assert matches[0].excel_source == "English"
    assert matches[0].excel_target == "Chinese (Simplified)"
    assert matches[0].rates == rates[("en", "zh-hans")]
