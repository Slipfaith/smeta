"""High-level import helpers for data ingestion from external sources."""

from __future__ import annotations

import os
from typing import Any, Dict, List, Tuple

from .outlook_import import map_message_to_project_info, parse_msg_file
from .outlook_import.msg_reader import OutlookMsgError
from .trados_xml_parser import parse_reports


def import_xml_reports(paths: List[str]) -> Tuple[Dict[str, Any], List[str]]:
    """Parse TRADOS XML reports and return structured information.

    Returns a tuple containing a dictionary with parsed data and a list of
    error messages.  The dictionary contains the parsed language pair volumes,
    report sources and any warnings emitted during parsing.
    """

    result: Dict[str, Any] = {"data": {}, "warnings": [], "report_sources": {}}
    errors: List[str] = []

    try:
        data, warnings, report_sources = parse_reports(paths)
    except Exception as exc:  # pragma: no cover - passthrough for GUI display
        errors.append(f"Ошибка при обработке XML файлов: {exc}")
        return result, errors

    result.update({"data": data, "warnings": warnings, "report_sources": report_sources})
    return result, errors


def import_project_info(paths: List[str]) -> Tuple[Dict[str, Any], List[str]]:
    """Extract project information from Outlook ``.msg`` files.

    The GUI provides a sequence of file paths.  The first successfully parsed
    message is converted into a serialisable structure that contains the
    project information along with metadata required to present results to the
    user.  If no files could be parsed, an empty result and a list of errors is
    returned.
    """

    errors: List[str] = []
    for path in paths:
        try:
            message = parse_msg_file(path)
            parse_result = map_message_to_project_info(message)
            return _prepare_project_info_payload(parse_result, path), []
        except (OutlookMsgError, RuntimeError) as exc:
            errors.append(f"{os.path.basename(path)}: {exc}")

    return {}, errors


def _prepare_project_info_payload(parse_result, source_path: str) -> Dict[str, Any]:
    """Convert ``ProjectInfoParseResult`` into a serialisable dictionary."""

    data = parse_result.data
    legal_entity_value = (data.legal_entity or "").strip()
    force_ru_mode = False
    if legal_entity_value.lower() == "logrus it usa":
        legal_entity_value = "Logrus IT"
        force_ru_mode = True

    payload: Dict[str, Any] = {
        "source_path": source_path,
        "data": {
            "project_name": data.project_name or "",
            "client_name": data.client_name or "",
            "contact_name": data.contact_name or "",
            "email": data.email or "",
            "legal_entity": legal_entity_value,
            "currency_code": data.currency_code or "",
            "raw_values": dict(data.raw_values),
        },
        "missing_fields": list(parse_result.missing_fields),
        "warnings": list(parse_result.warnings),
        "sender": {
            "name": parse_result.sender_name or "",
            "email": parse_result.sender_email or "",
            "sent_at": parse_result.sent_at or "",
        },
        "flags": {"force_ru_mode": force_ru_mode},
    }

    contact_email = getattr(data, "contact_email", None)
    if contact_email and not payload["data"].get("email"):
        payload["data"]["email"] = contact_email

    return payload
