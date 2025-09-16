import json
from json import JSONDecodeError
from pathlib import Path
from typing import Dict, Iterable, Tuple

from resource_utils import resource_path

CONFIG_RELATIVE_PATH = Path("logic") / "legal_entities.json"
CONFIG_PATH = resource_path(CONFIG_RELATIVE_PATH)

#: Built-in fallback mapping used when the JSON configuration is unavailable.
DEFAULT_LEGAL_ENTITIES: Dict[str, Path] = {
    "Артфест": Path("templates") / "Артфест.xlsx",
    "Бикрон": Path("templates") / "Бикрон.xlsx",
    "Логрус Айти": Path("templates") / "Логрус Айти.xlsx",
    "Logrus IT": Path("templates") / "Logrus IT.xlsx",
}


def _resolve_templates(items: Iterable[Tuple[str, Path | str]]) -> Dict[str, str]:
    """Convert relative template paths into absolute filesystem paths."""

    resolved: Dict[str, str] = {}
    for name, relative in items:
        resolved[name] = str(resource_path(Path(relative)))
    return resolved


def load_legal_entities() -> Dict[str, str]:
    """Return mapping of legal entity name to absolute template path."""

    try:
        with CONFIG_PATH.open("r", encoding="utf-8") as f:
            data = json.load(f)
    except (FileNotFoundError, JSONDecodeError):
        return _resolve_templates(DEFAULT_LEGAL_ENTITIES.items())

    if isinstance(data, dict):
        return _resolve_templates(data.items())

    return _resolve_templates(DEFAULT_LEGAL_ENTITIES.items())


def get_entities_list() -> Dict[str, str]:
    """Return mapping for convenience; kept for backward compatibility."""
    return load_legal_entities()
