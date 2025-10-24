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


def test_match_pairs_uses_manual_codes_without_normalization() -> None:
    rates = {("custom-src", "custom-tgt"): {"basic": 2.0, "complex": 3.0, "hour": 4.0}}

    manual_codes = {("Foo", "Bar"): ("custom-src", "custom-tgt")}
    manual_names = {("Foo", "Bar"): ("Manual Foo", "Manual Bar")}

    matches = rates_importer.match_pairs(
        [("Foo", "Bar")],
        rates,
        manual_codes=manual_codes,
        manual_names=manual_names,
    )

    assert matches[0].excel_source == "Manual Foo"
    assert matches[0].excel_target == "Manual Bar"
    assert matches[0].rates == rates[("custom-src", "custom-tgt")]
