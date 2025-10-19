import json
from json import JSONDecodeError
from pathlib import Path
from typing import Any, Dict, Iterable, Optional, Tuple

from resource_utils import resource_path
from .user_config import get_appdata_dir

DEFAULT_CONFIG_RELATIVE_PATH = Path("logic") / "legal_entities.json"
DEFAULT_CONFIG_PATH = resource_path(DEFAULT_CONFIG_RELATIVE_PATH)
USER_CONFIG_PATH = Path(get_appdata_dir()) / "legal_entities.json"
USER_TEMPLATES_DIR = Path(get_appdata_dir()) / "templates"
USER_LOGOS_DIR = USER_TEMPLATES_DIR / "logos"

_LEGAL_ENTITY_METADATA: Dict[str, Dict[str, Any]] = {}
_DEFAULT_ENTITY_NAMES: set[str] = set()


def _ensure_user_dirs() -> None:
    """Create directories for custom templates and logos if required."""

    USER_TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
    USER_LOGOS_DIR.mkdir(parents=True, exist_ok=True)


def _resolve_path(value: Path | str) -> str:
    """Return an absolute filesystem path for bundled or user files."""

    path = Path(value)
    if path.is_absolute():
        return str(path)
    return str(resource_path(path))


def _resolve_templates(items: Iterable[Tuple[str, Path | str]]) -> Dict[str, str]:
    """Convert stored template paths into absolute filesystem paths."""

    resolved: Dict[str, str] = {}
    for name, stored_path in items:
        resolved[name] = _resolve_path(Path(stored_path))
    return resolved


def _prepare_from_mapping(
    data: Dict[str, Any],
    user_entities: Optional[set[str]] = None,
) -> Dict[str, str]:
    templates: Dict[str, Path | str] = {}
    _LEGAL_ENTITY_METADATA.clear()
    user_entities = user_entities or set()

    for name, value in data.items():
        template: Optional[Path | str] = None
        metadata: Dict[str, Any] = {}
        if isinstance(value, dict):
            template = value.get("template")
            if value.get("logo"):
                metadata["logo"] = _resolve_path(Path(value["logo"]))
            for key in ("vat_enabled", "default_vat", "vat_rate", "display_name"):
                if key in value:
                    metadata[key] = value[key]
        elif isinstance(value, (str, Path)):
            template = value

        if not template:
            continue

        templates[name] = Path(template)
        resolved_template = _resolve_path(Path(template))
        metadata["template_path"] = resolved_template
        metadata["source"] = "user" if name in user_entities else "default"

        if "logo" not in metadata:
            metadata["logo"] = _guess_logo_path(name, resolved_template)
        _LEGAL_ENTITY_METADATA[name] = metadata

    return _resolve_templates(templates.items())


def _guess_logo_path(entity: str, template_path: str) -> Optional[str]:
    """Try to resolve a logo path for the given entity if it is not explicit."""

    candidates = []
    # Custom logos take priority.
    candidates.append(USER_LOGOS_DIR / f"{entity}.png")
    template_stem = Path(template_path).stem
    candidates.append(USER_LOGOS_DIR / f"{template_stem}.png")
    # Bundled defaults.
    candidates.append(resource_path(Path("templates") / "logos" / f"{entity}.png"))
    candidates.append(
        resource_path(Path("templates") / "logos" / f"{template_stem}.png")
    )

    for candidate in candidates:
        try:
            candidate_path = Path(candidate)
        except TypeError:
            continue
        if candidate_path.exists():
            return str(candidate_path)
    return None


def _load_config(path: Path) -> Dict[str, Any]:
    try:
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
            if isinstance(data, dict):
                return data
    except (FileNotFoundError, JSONDecodeError):
        return {}
    except Exception:
        return {}
    return {}


def _extract_entities(data: Dict[str, Any]) -> Dict[str, Any]:
    if "entities" in data and isinstance(data["entities"], dict):
        return data["entities"]
    if isinstance(data, dict):
        return data
    return {}


def _load_entities() -> Tuple[Dict[str, Any], Dict[str, Any]]:
    default_config = _extract_entities(_load_config(DEFAULT_CONFIG_PATH))
    user_config = _extract_entities(_load_config(USER_CONFIG_PATH))
    global _DEFAULT_ENTITY_NAMES
    _DEFAULT_ENTITY_NAMES = set(default_config.keys())
    return default_config, user_config


def load_legal_entities() -> Dict[str, str]:
    """Return mapping of legal entity name to absolute template path."""

    default_config, user_config = _load_entities()
    merged: Dict[str, Any] = dict(default_config)
    merged.update(user_config)
    return _prepare_from_mapping(merged, set(user_config.keys()))


def get_entities_list() -> Dict[str, str]:
    """Return mapping for convenience; kept for backward compatibility."""
    return load_legal_entities()


def get_legal_entity_metadata() -> Dict[str, Dict[str, Any]]:
    """Return metadata for legal entities loaded from configuration."""

    if not _LEGAL_ENTITY_METADATA:
        load_legal_entities()
    return {name: dict(meta) for name, meta in _LEGAL_ENTITY_METADATA.items()}


def save_user_entities(entities: Dict[str, Any]) -> bool:
    """Persist custom legal entities configuration to the user storage."""

    _ensure_user_dirs()
    try:
        with USER_CONFIG_PATH.open("w", encoding="utf-8") as f:
            json.dump({"entities": entities}, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def load_user_entities() -> Dict[str, Any]:
    """Return mapping of user-defined entities without defaults."""

    _, user_config = _load_entities()
    return dict(user_config)


def set_user_entity(name: str, value: Dict[str, Any]) -> bool:
    """Add or update a user-defined legal entity configuration."""

    user_entities = load_user_entities()
    user_entities[name] = value
    if save_user_entities(user_entities):
        load_legal_entities()
        return True
    return False


def remove_user_entity(name: str) -> bool:
    """Remove a user-defined legal entity (without touching bundled ones)."""

    user_entities = load_user_entities()
    if name not in user_entities:
        return False
    user_entities.pop(name, None)
    if save_user_entities(user_entities):
        load_legal_entities()
        return True
    return False


def get_user_templates_dir() -> Path:
    """Return the directory for storing custom templates."""

    _ensure_user_dirs()
    return USER_TEMPLATES_DIR


def get_user_logos_dir() -> Path:
    """Return the directory for storing custom logos."""

    _ensure_user_dirs()
    return USER_LOGOS_DIR


def is_default_entity(name: str) -> bool:
    """Check whether the entity originates from bundled configuration."""

    if not _DEFAULT_ENTITY_NAMES:
        load_legal_entities()
    return name in _DEFAULT_ENTITY_NAMES
