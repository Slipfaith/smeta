from copy import deepcopy
from typing import Any, Dict, List

from .settings_store import (
    get_additional_services,
    get_fuzzy_thresholds,
    get_translation_rows,
)


class ServiceConfig:
    """Конфигурация услуг и коэффициентов, загружаемая из пользовательских настроек."""

    TRANSLATION_ROWS: List[Dict[str, Any]] = []
    ADDITIONAL_SERVICES: Dict[str, List[Dict[str, Any]]] = {}
    ROW_NAMES: List[str] = []
    FUZZY_THRESHOLDS: Dict[str, int] = {}

    @classmethod
    def refresh(cls) -> None:
        cls.TRANSLATION_ROWS = get_translation_rows()
        cls.ADDITIONAL_SERVICES = get_additional_services()
        cls.ROW_NAMES = [row.get("name", "") for row in cls.TRANSLATION_ROWS]
        cls.FUZZY_THRESHOLDS = get_fuzzy_thresholds()

    @classmethod
    def copy_translation_rows(cls) -> List[Dict[str, Any]]:
        return deepcopy(cls.TRANSLATION_ROWS)

    @classmethod
    def copy_additional_services(cls) -> Dict[str, List[Dict[str, Any]]]:
        return deepcopy(cls.ADDITIONAL_SERVICES)


ServiceConfig.refresh()
