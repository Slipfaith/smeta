"""Compatibility wrappers around :mod:`logic.ms_graph_client`."""

from __future__ import annotations

import logging
from io import BytesIO
from typing import Iterable, Optional, Sequence, Union

import pandas as pd

from logic.ms_graph_client import (
    GraphClientError,
    authenticate_with_msal as _authenticate_with_msal,
    download_excel_from_file_id,
    download_excel_from_path,
)

logger = logging.getLogger(__name__)

ScopeInput = Union[str, Iterable[str], None]


def _normalise_scopes(scope: ScopeInput) -> Sequence[str]:
    if scope is None:
        return []
    if isinstance(scope, str):
        scope = scope.strip()
        return [scope] if scope else []
    try:
        scopes = list(scope)
    except TypeError:
        return [str(scope)]
    return [s for s in scopes if s]


def authenticate_with_msal(
    client_id: str,
    tenant_id: str,
    scope: ScopeInput,
) -> Optional[str]:
    """Return an access token for ``scope`` using MSAL.

    This helper mirrors the API provided by the reference ``rates1`` package.
    Any errors are logged and ``None`` is returned so the legacy UI can fail
    gracefully without crashing the whole application.
    """

    scopes = _normalise_scopes(scope)
    try:
        return _authenticate_with_msal(client_id, tenant_id, scopes)
    except GraphClientError as exc:
        logger.error("MSAL authentication failed: %s", exc)
    except Exception as exc:  # pragma: no cover - unexpected MSAL errors
        logger.exception("Unexpected MSAL error: %s", exc)
    return None


def _read_excel(payload: bytes, **read_kwargs) -> Optional[pd.DataFrame]:
    if not payload:
        return None
    try:
        return pd.read_excel(BytesIO(payload), **read_kwargs)
    except Exception as exc:  # pragma: no cover - pandas/openpyxl errors are rare
        logger.error("Failed to read Excel payload: %s", exc)
        return None


def download_excel_from_sharepoint(
    access_token: str,
    site_id: str,
    file_path: str,
) -> Optional[pd.DataFrame]:
    """Download and decode an Excel workbook by ``file_path``.

    Returns a :class:`pandas.DataFrame` on success or ``None`` if anything goes
    wrong.  The signature matches the one used by the legacy UI.
    """

    if not access_token:
        return None

    try:
        payload = download_excel_from_path(access_token, site_id, file_path)
    except GraphClientError as exc:
        logger.error("Failed to download Excel file from SharePoint: %s", exc)
        return None

    return _read_excel(payload)


def download_excel_by_fileid(
    access_token: str,
    site_id: str,
    file_id: str,
    sheet_name: Optional[str] = None,
    skiprows: int = 0,
) -> Optional[pd.DataFrame]:
    """Download and decode an Excel workbook identified by ``file_id``."""

    if not access_token:
        return None

    try:
        payload = download_excel_from_file_id(access_token, site_id, file_id)
    except GraphClientError as exc:
        logger.error("Failed to download Excel file by id: %s", exc)
        return None

    read_kwargs = {"skiprows": skiprows} if skiprows else {}
    if sheet_name is not None:
        read_kwargs["sheet_name"] = sheet_name
        read_kwargs.setdefault("header", None)
    elif skiprows:
        read_kwargs.setdefault("header", None)

    return _read_excel(payload, **read_kwargs)
