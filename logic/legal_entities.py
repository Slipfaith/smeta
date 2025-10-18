import json
import shutil
from dataclasses import dataclass
from json import JSONDecodeError
from pathlib import Path
from typing import Any, Dict, Iterable, Optional, Tuple

from resource_utils import resource_path

from .user_config import get_appdata_dir

CONFIG_RELATIVE_PATH = Path("logic") / "legal_entities.json"
CONFIG_PATH = resource_path(CONFIG_RELATIVE_PATH)
USER_CONFIG_PATH = Path(get_appdata_dir()) / "legal_entities.json"
USER_TEMPLATES_DIR = Path(get_appdata_dir()) / "legal_entity_templates"
USER_LOGOS_DIR = Path(get_appdata_dir()) / "legal_entity_logos"

USER_TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
USER_LOGOS_DIR.mkdir(parents=True, exist_ok=True)


@dataclass
class LegalEntityRecord:
    name: str
    template: str
    metadata: Dict[str, Any]
    source: str = "built-in"


_LEGAL_ENTITY_METADATA: Dict[str, Dict[str, Any]] = {}

def _resolve_templates(items: Iterable[Tuple[str, Path | str]]) -> Dict[str, str]:
    """Convert relative template paths into absolute filesystem paths."""

    resolved: Dict[str, str] = {}
    for name, relative in items:
        resolved[name] = str(resource_path(Path(relative)))
    return resolved


def _store_logo_file(name: str, source_path: Path) -> str:
    USER_LOGOS_DIR.mkdir(parents=True, exist_ok=True)
    resolved_source = source_path.resolve()
    target_dir = USER_LOGOS_DIR.resolve()
    try:
        is_internal = resolved_source.is_relative_to(target_dir)  # type: ignore[attr-defined]
    except AttributeError:  # pragma: no cover - Python < 3.9 fallback
        is_internal = target_dir in resolved_source.parents
    if is_internal:
        return str(resolved_source)
    target_path = target_dir / f"{name}{resolved_source.suffix or '.png'}"
    try:
        shutil.copy2(resolved_source, target_path)
    except Exception as exc:  # pragma: no cover - filesystem dependent
        raise IOError(f"Не удалось сохранить логотип: {exc}") from exc
    return str(target_path)


def _remove_logo_file(path: str) -> None:
    try:
        logo_path = Path(path)
        resolved = logo_path.resolve()
        target_dir = USER_LOGOS_DIR.resolve()
        try:
            is_internal = resolved.is_relative_to(target_dir)  # type: ignore[attr-defined]
        except AttributeError:  # pragma: no cover - Python < 3.9 fallback
            is_internal = target_dir in resolved.parents
        if is_internal:
            logo_path.unlink(missing_ok=True)
    except Exception:  # pragma: no cover - best effort cleanup
        pass


def _prepare_from_mapping(data: Dict[str, Any], source: str) -> Dict[str, LegalEntityRecord]:
    records: Dict[str, LegalEntityRecord] = {}

    for name, value in data.items():
        if isinstance(value, dict):
            template = value.get("template")
            metadata = {k: v for k, v in value.items() if k != "template"}
        else:
            template = value
            metadata = {}
        if not template:
            continue
        resolved = _resolve_templates([(name, template)]).get(name)
        if resolved is None:
            continue
        if source == "built-in" and isinstance(metadata, dict):
            logo_value = metadata.get("logo")
            if isinstance(logo_value, str) and logo_value:
                metadata["logo"] = str(resource_path(Path(logo_value)))
        record = LegalEntityRecord(
            name=name,
            template=str(resolved),
            metadata=metadata,
            source=source,
        )
        records[name] = record
    return records


def _load_json(path: Path) -> Dict[str, Any]:
    try:
        with path.open("r", encoding="utf-8") as fh:
            data = json.load(fh)
    except (FileNotFoundError, JSONDecodeError):
        return {}
    if isinstance(data, dict):
        if "entities" in data and isinstance(data["entities"], dict):
            return data["entities"]
        return data
    return {}


def _collect_records() -> Dict[str, LegalEntityRecord]:
    records: Dict[str, LegalEntityRecord] = {}
    _LEGAL_ENTITY_METADATA.clear()

    for mapping, source in ((CONFIG_PATH, "built-in"), (USER_CONFIG_PATH, "user")):
        data = _load_json(mapping)
        if not data:
            continue
        prepared = _prepare_from_mapping(data, source)
        records.update(prepared)
        for name, record in prepared.items():
            if record.metadata:
                _LEGAL_ENTITY_METADATA[name] = dict(record.metadata)

    return records


def load_legal_entities() -> Dict[str, str]:
    """Return mapping of legal entity name to absolute template path."""

    records = _collect_records()
    return {name: record.template for name, record in records.items()}


def get_entities_list() -> Dict[str, str]:
    """Return mapping for convenience; kept for backward compatibility."""
    return load_legal_entities()


def get_legal_entity_metadata() -> Dict[str, Dict[str, Any]]:
    """Return metadata for legal entities loaded from configuration."""

    if not _LEGAL_ENTITY_METADATA:
        load_legal_entities()
    return {name: dict(meta) for name, meta in _LEGAL_ENTITY_METADATA.items()}


def list_legal_entities_with_sources() -> Dict[str, LegalEntityRecord]:
    """Return legal entities mapped to detailed records including the source."""

    return _collect_records()


def load_user_legal_entities() -> Dict[str, Any]:
    return _load_json(USER_CONFIG_PATH)


def save_user_legal_entities(data: Dict[str, Any]) -> None:
    USER_CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with USER_CONFIG_PATH.open("w", encoding="utf-8") as fh:
        json.dump({"entities": data}, fh, ensure_ascii=False, indent=2)


def add_or_update_legal_entity(name: str, template_path: Path, metadata: Dict[str, Any] | None = None) -> LegalEntityRecord:
    name = name.strip()
    if not name:
        raise ValueError("Имя юридического лица не может быть пустым")
    if not template_path.exists():
        raise FileNotFoundError(template_path)

    USER_TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
    resolved_source = template_path.resolve()
    resolved_target_dir = USER_TEMPLATES_DIR.resolve()
    try:
        is_internal = resolved_source.is_relative_to(resolved_target_dir)  # type: ignore[attr-defined]
    except AttributeError:  # pragma: no cover - Python < 3.9 fallback
        is_internal = resolved_target_dir in resolved_source.parents
    if is_internal:
        target_path = resolved_source
    else:
        target_path = resolved_target_dir / template_path.name
        try:
            shutil.copy2(resolved_source, target_path)
        except Exception as exc:  # pragma: no cover - OS specific
            raise IOError(f"Не удалось сохранить шаблон: {exc}") from exc

    user_data = load_user_legal_entities()
    existing_entry = user_data.get(name, {}) if isinstance(user_data, dict) else {}
    metadata = dict(metadata or {})
    stored_logo = existing_entry.get("logo") if isinstance(existing_entry, dict) else None
    if "logo" in metadata:
        logo_value = metadata.get("logo")
        if isinstance(logo_value, str) and logo_value:
            stored_logo = _store_logo_file(name, Path(logo_value))
            existing_logo = existing_entry.get("logo") if isinstance(existing_entry, dict) else None
            if existing_logo and stored_logo != existing_logo:
                _remove_logo_file(existing_logo)
        else:
            if stored_logo:
                _remove_logo_file(stored_logo)
            stored_logo = None
    elif stored_logo:
        metadata["logo"] = stored_logo
    if stored_logo:
        metadata["logo"] = stored_logo
    else:
        metadata.pop("logo", None)

    payload = {"template": str(target_path)} | metadata
    user_data[name] = payload
    save_user_legal_entities(user_data)
    _collect_records()
    return LegalEntityRecord(name=name, template=str(target_path), metadata=metadata, source="user")


def remove_user_legal_entity(name: str) -> None:
    user_data = load_user_legal_entities()
    if name in user_data:
        template = user_data[name].get("template")
        logo = user_data[name].get("logo")
        user_data.pop(name, None)
        save_user_legal_entities(user_data)
        if template:
            try:
                path = Path(template)
                if path.is_file() and USER_TEMPLATES_DIR in path.parents:
                    path.unlink(missing_ok=True)
            except Exception:  # pragma: no cover - best effort cleanup
                pass
        if logo:
            _remove_logo_file(str(logo))
        _collect_records()


def export_legal_entities_to_excel(path: Path, records: Iterable[LegalEntityRecord] | None = None) -> None:
    from openpyxl import Workbook

    if records is None:
        records = list(list_legal_entities_with_sources().values())
    else:
        records = list(records)
    wb = Workbook()
    ws = wb.active
    ws.title = "LegalEntities"
    ws.append([
        "Название",
        "Путь к шаблону",
        "Источник",
        "VAT включён",
        "НДС по умолчанию",
    ])
    for record in records:
        metadata = record.metadata or {}
        ws.append([
            record.name,
            record.template,
            record.source,
            metadata.get("vat_enabled", ""),
            metadata.get("default_vat", ""),
        ])
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def get_legal_entity_logo(name: str) -> Optional[str]:
    metadata = get_legal_entity_metadata().get(name, {})
    logo = metadata.get("logo") if isinstance(metadata, dict) else None
    if isinstance(logo, str) and logo:
        logo_path = Path(logo)
        if logo_path.exists():
            return str(logo_path)
    fallback = resource_path(Path("templates") / "logos" / f"{name}.png")
    if fallback.exists():
        return str(fallback)
    return None
