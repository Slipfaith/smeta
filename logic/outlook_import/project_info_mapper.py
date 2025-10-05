"""Map Outlook message content to project information fields."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from .msg_reader import OutlookMessage
from .table_parser import extract_first_table

_PROJECT_NAME_RE = re.compile(r"\[(?P<name>[^\[\]]+)\]")

_KEY_NORMALIZATION = {
    "название клиента": "client_name",
    "company's name": "client_name",
    "company’s name": "client_name",
    "контрагент logrus it": "legal_entity",
    "контрагент logrus it (с ндс или нет, если с ндс, укажите размер ндс)": "legal_entity",
    "legal entity": "legal_entity",
    "контактное лицо со стороны клиента": "contact_person",
    "contact person": "contact_person",
    "email": "email",
    "email адрес": "email",
    "валюта расчетов": "currency",
    "currency": "currency",
}


@dataclass
class ProjectInfoData:
    project_name: Optional[str] = None
    client_name: Optional[str] = None
    contact_name: Optional[str] = None
    contact_email: Optional[str] = None
    email: Optional[str] = None
    legal_entity: Optional[str] = None
    currency_code: Optional[str] = None
    raw_values: Dict[str, str] = field(default_factory=dict)


@dataclass
class ProjectInfoParseResult:
    data: ProjectInfoData
    missing_fields: List[str]
    sender_name: Optional[str]
    sender_email: Optional[str]
    sent_at: Optional[str]
    warnings: List[str] = field(default_factory=list)


def _normalize_key(value: str) -> str:
    value = value.strip().lower()
    value = value.replace("’", "'")
    value = re.sub(r"\s+", " ", value)
    return _KEY_NORMALIZATION.get(value, value)


def _split_contact_value(value: str) -> (Optional[str], Optional[str]):
    value = value.strip()
    if not value:
        return None, None
    emails = re.findall(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", value)
    email = emails[0] if emails else None
    if email:
        remainder = value.replace(email, " ")
    else:
        remainder = value
    remainder = remainder.replace(",", " ").replace(";", " ")
    name = re.sub(r"\s+", " ", remainder).strip()
    if name and email and name.lower() == email.lower():
        name = None
    return name or None, email


def _normalize_currency(value: str) -> Optional[str]:
    value = value.strip().lower()
    if not value:
        return None
    replacements = {
        "руб": "RUB",
        "rub": "RUB",
        "rubbles": "RUB",
        "российский рубль": "RUB",
        "дол": "USD",
        "usd": "USD",
        "us dollar": "USD",
        "евр": "EUR",
        "eur": "EUR",
        "euro": "EUR",
    }
    for key, code in replacements.items():
        if key in value:
            return code
    value = value.upper()
    if value in {"RUB", "USD", "EUR"}:
        return value
    return None


def map_message_to_project_info(message: OutlookMessage) -> ProjectInfoParseResult:
    table_rows: List[List[str]] = []
    warnings: List[str] = []

    if message.html_body:
        table_rows = extract_first_table(message.html_body)
    if not table_rows and message.body:
        table_rows = extract_first_table(message.body)

    mapped_values: Dict[str, str] = {}
    if table_rows:
        for row in table_rows:
            if len(row) < 2:
                continue
            key = _normalize_key(row[0])
            value = row[1].strip()
            if not key or not value:
                continue
            mapped_values[key] = value
    else:
        warnings.append("Таблица в письме не найдена")

    project_name = None
    match = _PROJECT_NAME_RE.search(message.subject)
    if match:
        project_name = match.group("name").strip()

    client_name = mapped_values.get("client_name")

    contact_name = None
    contact_email = None
    contact_raw = mapped_values.get("contact_person")
    if contact_raw:
        contact_name, contact_email = _split_contact_value(contact_raw)
    email_value = mapped_values.get("email")
    maybe_email = None
    if email_value:
        email_value = email_value.strip()
        # if email row contains name, try to split
        maybe_name, maybe_email = _split_contact_value(email_value)
        if maybe_email:
            contact_email = contact_email or maybe_email
        if maybe_name and not contact_name:
            contact_name = maybe_name
    email = contact_email or maybe_email
    if not email and email_value and re.search(r"@", email_value):
        email = email_value

    legal_entity = mapped_values.get("legal_entity")
    currency_code = None
    currency_value = mapped_values.get("currency")
    if currency_value:
        currency_code = _normalize_currency(currency_value)
        if currency_code is None:
            warnings.append(f"Не удалось распознать валюту: {currency_value}")

    data = ProjectInfoData(
        project_name=project_name,
        client_name=client_name,
        contact_name=contact_name,
        contact_email=contact_email,
        email=email,
        legal_entity=legal_entity,
        currency_code=currency_code,
        raw_values=mapped_values,
    )

    missing = []
    if not project_name:
        missing.append("Название проекта")
    if not client_name:
        missing.append("Название клиента")
    if not (contact_name or contact_email):
        missing.append("Контактное лицо")
    if not (email):
        missing.append("Email")
    if not legal_entity:
        missing.append("Юрлицо")
    if not currency_code:
        missing.append("Валюта")

    sent_at = None
    if message.sent_at:
        sent_at = message.sent_at.isoformat(sep=" ", timespec="minutes")

    return ProjectInfoParseResult(
        data=data,
        missing_fields=missing,
        sender_name=message.sender_name,
        sender_email=message.sender_email,
        sent_at=sent_at,
        warnings=warnings,
    )
