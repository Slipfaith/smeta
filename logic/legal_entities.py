import json
from json import JSONDecodeError
from pathlib import Path
from typing import Any, Dict, Iterable, Tuple

from resource_utils import resource_path

CONFIG_RELATIVE_PATH = Path("logic") / "legal_entities.json"
CONFIG_PATH = resource_path(CONFIG_RELATIVE_PATH)

_LEGAL_ENTITY_METADATA: Dict[str, Dict[str, Any]] = {}

def _resolve_templates(items: Iterable[Tuple[str, Path | str]]) -> Dict[str, str]:
    """Convert relative template paths into absolute filesystem paths."""

    resolved: Dict[str, str] = {}
    for name, relative in items:
        resolved[name] = str(resource_path(Path(relative)))
    return resolved


def _prepare_from_mapping(data: Dict[str, Any]) -> Dict[str, str]:
    templates: Dict[str, Path | str] = {}
    _LEGAL_ENTITY_METADATA.clear()

    for name, value in data.items():
        if isinstance(value, dict):
            template = value.get("template")
            if template:
                templates[name] = Path(template)
            metadata = {k: v for k, v in value.items() if k != "template"}
            if metadata:
                _LEGAL_ENTITY_METADATA[name] = metadata
        elif isinstance(value, (str, Path)):
            templates[name] = Path(value)

    return _resolve_templates(templates.items())


def load_legal_entities() -> Dict[str, str]:
    """Return mapping of legal entity name to absolute template path."""

    try:
        with CONFIG_PATH.open("r", encoding="utf-8") as f:
            data = json.load(f)
    except (FileNotFoundError, JSONDecodeError):
        _LEGAL_ENTITY_METADATA.clear()
        return {}

    if isinstance(data, dict):
        if "entities" in data and isinstance(data["entities"], dict):
            return _prepare_from_mapping(data["entities"])
        return _prepare_from_mapping(data)

    return {}


def get_entities_list() -> Dict[str, str]:
    """Return mapping for convenience; kept for backward compatibility."""
    return load_legal_entities()


def get_legal_entity_metadata() -> Dict[str, Dict[str, Any]]:
    """Return metadata for legal entities loaded from configuration."""

    if not _LEGAL_ENTITY_METADATA:
        load_legal_entities()
    return {name: dict(meta) for name, meta in _LEGAL_ENTITY_METADATA.items()}
