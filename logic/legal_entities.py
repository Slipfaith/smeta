import json
import os
from typing import Dict

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "legal_entities.json")


def load_legal_entities() -> Dict[str, str]:
    """Return mapping of legal entity name to absolute template path."""
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)
    base_dir = os.path.normpath(os.path.join(os.path.dirname(__file__), ".."))
    out: Dict[str, str] = {}
    if isinstance(data, dict):
        for name, rel_path in data.items():
            out[name] = os.path.join(base_dir, rel_path)
    return out


def get_entities_list() -> Dict[str, str]:
    """Return mapping for convenience; kept for backward compatibility."""
    return load_legal_entities()
