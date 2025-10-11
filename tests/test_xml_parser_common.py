from logic.xml_parser_common import language_identity


def test_language_identity_handles_override_aliases() -> None:
    assert language_identity("Испанский (Латам)") == ("es", "", "419")
