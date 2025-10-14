from logic.xml_parser_common import expand_target_matches


def test_expand_target_matches_includes_all_variants():
    matched = {"es"}
    available = {
        "es": "Spanish",
        "es-es": "Spanish (Spain)",
        "es-419": "Spanish (Latam)",
    }

    result = expand_target_matches(matched, available.keys(), "en")

    assert result == {"es", "es-es", "es-419"}


def test_expand_target_matches_excludes_source_language():
    matched = {"en-us"}
    available = {
        "en-us": "English (United States)",
        "en-gb": "English (United Kingdom)",
    }

    result = expand_target_matches(matched, available.keys(), "en-us")

    assert "en-us" not in result
    assert "en-gb" in result
