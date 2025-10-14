from logic.xml_parser_common import language_identity


def test_language_identity_handles_override_aliases() -> None:
    assert language_identity("Испанский (Латам)") == ("es", "", "419")


def test_language_identity_prefers_full_language_name_over_partial_code() -> None:
    # ``Malay`` should not be mistaken for the ISO-639 code ``mal`` (Malayalam).
    assert language_identity("Malay") == ("ms", "", "")
