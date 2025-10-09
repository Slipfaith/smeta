import pytest

from logic.translation_config import tr


def test_tr_returns_translation_when_available():
    assert tr("Перевод", "en") == "Translation"


def test_tr_falls_back_to_original_when_missing_phrase():
    missing = "Unknown phrase"
    assert tr(missing, "ru") == missing


def test_tr_falls_back_to_original_when_missing_language():
    assert tr("Перевод", "de") == "Перевод"


@pytest.mark.parametrize(
    ("phrase", "expected"),
    [
        ("Скидка", "Discount"),
        ("Наценка", "Markup"),
        ("Сумма скидки", "Discount amount"),
        ("Сумма наценки", "Markup amount"),
        ("Сумма со скидкой", "Total with discount"),
    ],
)
def test_discount_and_markup_phrases_have_english_translations(phrase, expected):
    assert tr(phrase, "en") == expected
