"""Reading data from Outlook .msg files."""

from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass
from typing import Optional

try:  # pragma: no cover - import tested indirectly
    import extract_msg  # type: ignore
except ImportError:  # pragma: no cover - dependency resolved at runtime
    extract_msg = None  # type: ignore


class OutlookMsgError(Exception):
    """Raised when an Outlook .msg file cannot be parsed."""


def _count_cyrillic(text: str) -> int:
    return sum(
        "А" <= ch <= "я" or ch in {"Ё", "ё"}
        for ch in text
    )


def normalize_outlook_text(value: str) -> str:
    """Normalize Outlook strings, fixing common CP1251 mojibake."""

    if not value:
        return value

    baseline = _count_cyrillic(value)
    if baseline:
        return value

    try:
        candidate = value.encode("latin-1", errors="strict").decode("cp1251")
    except UnicodeEncodeError:
        return value

    candidate_cyrillic = _count_cyrillic(candidate)
    if candidate_cyrillic > baseline and candidate_cyrillic > 0:
        return candidate
    return value


@dataclass
class OutlookMessage:
    """Normalized representation of an Outlook message."""

    subject: str
    sender_name: Optional[str]
    sender_email: Optional[str]
    sent_at: Optional[_dt.datetime]
    body: str
    html_body: Optional[str]

    def __post_init__(self) -> None:
        self.subject = normalize_outlook_text((self.subject or "").strip())


def _parse_datetime(value: Optional[str]) -> Optional[_dt.datetime]:
    if not value:
        return None
    # ``extract_msg`` typically returns an RFC 2822 like string or ISO.
    # Attempt a couple of common formats.
    for fmt in (
        "%a, %d %b %Y %H:%M:%S %z",
        "%a, %d %b %Y %H:%M:%S",
        "%d.%m.%Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S%z",
        "%Y-%m-%d %H:%M:%S",
    ):
        try:
            return _dt.datetime.strptime(value, fmt)
        except ValueError:
            continue
    return None


def parse_msg_file(path: str) -> OutlookMessage:
    """Parse ``path`` into an :class:`OutlookMessage` instance."""

    if extract_msg is None:
        raise OutlookMsgError(
            "Package 'extract_msg' is required to parse Outlook .msg files"
        )

    try:
        msg = extract_msg.Message(path)
    except Exception as exc:  # pragma: no cover - library level errors
        raise OutlookMsgError(str(exc)) from exc

    try:
        subject = normalize_outlook_text((msg.subject or "").strip())
        body = msg.body or ""
        html_body = getattr(msg, "htmlBody", None) or None

        sender_name = getattr(msg, "sender", None) or getattr(msg, "sender_name", None)
        if isinstance(sender_name, str):
            sender_name = normalize_outlook_text(sender_name.strip()) or None
        sender_email = getattr(msg, "sender_email", None) or getattr(
            msg, "sender_email_address", None
        )

        # ``extract_msg`` stores delivery time on ``date`` or ``receivedTime``
        sent_raw = getattr(msg, "date", None) or getattr(msg, "receivedTime", None)
        sent_at = _parse_datetime(sent_raw if isinstance(sent_raw, str) else None)
    finally:
        msg.close()

    return OutlookMessage(
        subject=subject.strip(),
        sender_name=(sender_name or None),
        sender_email=(sender_email or None),
        sent_at=sent_at,
        body=body,
        html_body=html_body,
    )
