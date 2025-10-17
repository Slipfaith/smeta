import json
import shutil
from json import JSONDecodeError
from pathlib import Path
from typing import Any, Dict, Iterable, List, Tuple

from resource_utils import resource_path
from logic.user_config import get_appdata_dir

CONFIG_RELATIVE_PATH = Path("logic") / "legal_entities.json"
CONFIG_PATH = resource_path(CONFIG_RELATIVE_PATH)
USER_CONFIG_PATH = Path(get_appdata_dir()) / "legal_entities.json"
USER_TEMPLATES_DIR = Path(get_appdata_dir()) / "legal_entity_templates"

_LEGAL_ENTITY_METADATA: Dict[str, Dict[str, Any]] = {}


def _load_raw_mapping(path: Path) -> Dict[str, Any]:
    try:
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
    except (FileNotFoundError, JSONDecodeError):
        return {}
    if isinstance(data, dict):
        if "entities" in data and isinstance(data["entities"], dict):
            return dict(data["entities"])
        return dict(data)
    return {}


def _resolve_template_path(value: Path | str) -> str:
    path = Path(value)
    if path.is_absolute():
        return str(path)
    return str(resource_path(path))


def _extract_template_entry(value: Any) -> Tuple[Path | str | None, Dict[str, Any]]:
    if isinstance(value, dict):
        template = value.get("template")
        metadata = {k: v for k, v in value.items() if k != "template"}
        return (template, metadata)
    if isinstance(value, (str, Path)):
        return (value, {})
    return (None, {})


def _prepare_from_mapping(data: Dict[str, Any]) -> Tuple[Dict[str, str], Dict[str, Dict[str, Any]]]:
    templates: Dict[str, str] = {}
    metadata: Dict[str, Dict[str, Any]] = {}
    for name, value in data.items():
        template, extra = _extract_template_entry(value)
        if template:
            templates[name] = _resolve_template_path(Path(template))
        if extra:
            metadata[name] = extra
    return templates, metadata


def _merge_configs(configs: Iterable[Dict[str, Any]]) -> Tuple[Dict[str, str], Dict[str, Dict[str, Any]]]:
    merged_templates: Dict[str, str] = {}
    merged_metadata: Dict[str, Dict[str, Any]] = {}
    for cfg in configs:
        templates, metadata = _prepare_from_mapping(cfg)
        merged_templates.update(templates)
        merged_metadata.update(metadata)
    return merged_templates, merged_metadata


def load_user_legal_entities() -> Dict[str, Any]:
    return _load_raw_mapping(USER_CONFIG_PATH)


def save_user_legal_entities(data: Dict[str, Any]) -> None:
    USER_CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with USER_CONFIG_PATH.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def add_or_update_legal_entity(name: str, template_path: str, metadata: Dict[str, Any] | None = None) -> None:
    current = load_user_legal_entities()
    entry: Dict[str, Any] = {"template": template_path}
    if metadata:
        entry.update(metadata)
    current[name] = entry
    save_user_legal_entities(current)


def remove_user_legal_entity(name: str) -> None:
    current = load_user_legal_entities()
    if name in current:
        current.pop(name)
        save_user_legal_entities(current)


def ensure_user_template_copy(source: Path, name: str) -> Path:
    USER_TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
    safe_name = name.replace("/", "_").replace("\\", "_")
    destination = USER_TEMPLATES_DIR / (safe_name + source.suffix)
    counter = 1
    while destination.exists():
        destination = USER_TEMPLATES_DIR / f"{safe_name}_{counter}{source.suffix}"
        counter += 1
    shutil.copy2(source, destination)
    return destination


def load_legal_entities() -> Dict[str, str]:
    """Return mapping of legal entity name to absolute template path."""

    base = _load_raw_mapping(CONFIG_PATH)
    user = load_user_legal_entities()
    templates, metadata = _merge_configs([base, user])
    _LEGAL_ENTITY_METADATA.clear()
    _LEGAL_ENTITY_METADATA.update(metadata)
    return templates


def get_entities_list() -> Dict[str, str]:
    """Return mapping for convenience; kept for backward compatibility."""
    return load_legal_entities()


def get_legal_entity_metadata() -> Dict[str, Dict[str, Any]]:
    """Return metadata for legal entities loaded from configuration."""

    if not _LEGAL_ENTITY_METADATA:
        load_legal_entities()
    return {name: dict(meta) for name, meta in _LEGAL_ENTITY_METADATA.items()}


def list_legal_entities_detailed() -> List[Dict[str, Any]]:
    """Return detailed information about legal entities for settings UI."""

    base = _load_raw_mapping(CONFIG_PATH)
    user = load_user_legal_entities()
    details: List[Dict[str, Any]] = []
    names = set(base.keys()) | set(user.keys())
    for name in sorted(names, key=str.casefold):
        if name in user:
            source = "user"
            raw = user[name]
        else:
            source = "default"
            raw = base.get(name)
        template, metadata = _extract_template_entry(raw)
        resolved = _resolve_template_path(Path(template)) if template else ""
        details.append(
            {
                "name": name,
                "template": str(template) if template else "",
                "resolved_template": resolved,
                "metadata": metadata,
                "source": source,
            }
        )
    return details


def export_legal_entity_template(name: str, destination: Path) -> bool:
    entries = {entry["name"]: entry for entry in list_legal_entities_detailed()}
    entry = entries.get(name)
    if not entry:
        return False
    src = Path(entry.get("resolved_template") or "")
    if not src.exists():
        return False
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, destination)
    return True
