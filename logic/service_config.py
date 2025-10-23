class ServiceConfig:
    """Конфигурация услуг и коэффициентов."""

    # ---- Услуги перевода (без редактирования) ----
    TRANSLATION_ROWS = [
        {"key": "new", "name": "Перевод, новые слова (100%)", "multiplier": 1.0, "is_base": True},
        {"key": "fuzzy_75_94", "name": "Перевод, совпадения 75-94% (66%)", "multiplier": 0.66, "is_base": False},
        {"key": "fuzzy_95_99", "name": "Перевод, совпадения 95-99% (33%)", "multiplier": 0.33, "is_base": False},
        {"key": "reps_100_30", "name": "Перевод, повторы и 100% совпадения (30%)", "multiplier": 0.30, "is_base": False},
    ]

    # Имена строк статистики, используемые при экспорте и парсинге отчётов
    ROW_NAMES = [row["name"] for row in TRANSLATION_ROWS]

