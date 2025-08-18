class ServiceConfig:
    """Конфигурация услуг и их параметров"""

    TRANSLATION_ROWS = [
        {"name": "Перевод, новые слова (100%)", "multiplier": 1.0, "is_base": True},
        {"name": "Перевод, совпадения 75-94% (66%)", "multiplier": 0.66, "is_base": False},
        {"name": "Перевод, совпадения 95-99% (33%)", "multiplier": 0.33, "is_base": False},
        {"name": "Перевод, повторы и 100% совпадения (30%)", "multiplier": 0.30, "is_base": False}
    ]

    EDITING_ROWS = [
        {"name": "Редактирование новых слов", "multiplier": 1.0, "is_base": True},
        {"name": "Редактирование повторов", "multiplier": 0.5, "is_base": False},
        {"name": "Корректорская правка", "multiplier": 1.0, "is_base": True},
        {"name": "Стилистическая правка", "multiplier": 1.0, "is_base": True}
    ]

    ADDITIONAL_SERVICES = {
        "Верстка": [
            {"name": "InDesign верстка", "multiplier": 1.0, "is_base": True},
            {"name": "PowerPoint верстка", "multiplier": 1.0, "is_base": True},
            {"name": "PDF верстка", "multiplier": 1.0, "is_base": True},
            {"name": "Графика/Изображения", "multiplier": 1.0, "is_base": True}
        ],
        "Локализация мультимедиа": [
            {"name": "Создание субтитров", "multiplier": 1.0, "is_base": True},
            {"name": "Озвучка", "multiplier": 1.0, "is_base": True},
            {"name": "Видеомонтаж", "multiplier": 1.0, "is_base": True},
            {"name": "Синхронизация", "multiplier": 1.0, "is_base": True}
        ],
        "Тестирование/QA": [
            {"name": "Лингвистическое тестирование", "multiplier": 1.0, "is_base": True},
            {"name": "Функциональное тестирование", "multiplier": 1.0, "is_base": True},
            {"name": "Косметическое тестирование", "multiplier": 1.0, "is_base": True},
            {"name": "Финальная проверка", "multiplier": 1.0, "is_base": True}
        ],
        "Прочие услуги": [
            {"name": "Создание терминологии", "multiplier": 1.0, "is_base": True},
            {"name": "Подготовка Translation Memory", "multiplier": 1.0, "is_base": True},
            {"name": "Анализ CAT-инструмента", "multiplier": 1.0, "is_base": True},
            {"name": "Консультации", "multiplier": 1.0, "is_base": True}
        ]
    }
