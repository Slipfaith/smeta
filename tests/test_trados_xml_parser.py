"""Tests for helpers used when parsing Trados analyse reports."""

from logic.trados_xml_parser import _extract_languages_from_filename


def test_extract_languages_supports_numeric_region_codes() -> None:
    """Ensure filenames with numeric UN M.49 region codes are handled."""

    src, tgt = _extract_languages_from_filename(
        "Analyze Files en-US_es-419(42).xml"
    )

    assert src == "Английский (Соединенные Штаты)"
    assert tgt == "Испанский (Латам)"
