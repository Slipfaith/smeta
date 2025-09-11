import json
from pathlib import Path
from typing import Dict

from resource_utils import resource_path

CONFIG_PATH = resource_path(Path("logic") / "legal_entities.json")


def load_legal_entities() -> Dict[str, str]:
    """Return mapping of legal entity name to absolute template path."""
    with CONFIG_PATH.open("r", encoding="utf-8") as f:
        data = json.load(f)
    out: Dict[str, str] = {}
    if isinstance(data, dict):
        for name, rel_path in data.items():
            out[name] = str(resource_path(rel_path))
    return out


def get_entities_list() -> Dict[str, str]:
    """Return mapping for convenience; kept for backward compatibility."""
    return load_legal_entities()
