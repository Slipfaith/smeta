"""Helpers for working with Microsoft Graph."""

from __future__ import annotations

import logging
from typing import Iterable

import msal
import requests

logger = logging.getLogger(__name__)


class GraphClientError(RuntimeError):
    """Raised when communication with Microsoft Graph fails."""


def authenticate_with_msal(
    client_id: str,
    tenant_id: str,
    scopes: Iterable[str],
) -> str:
    """Return an access token for the provided *scopes* using MSAL."""

    if not client_id or not tenant_id:
        raise GraphClientError("Missing client or tenant identifier for MSAL authentication")

    scopes = list(scopes)
    if not scopes:
        raise GraphClientError("No OAuth scopes configured for MSAL authentication")

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.PublicClientApplication(client_id, authority=authority)

    try:
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(scopes, account=accounts[0])
        else:
            result = app.acquire_token_interactive(scopes=scopes)
    except Exception as exc:  # pragma: no cover - MSAL specific error handling
        raise GraphClientError(f"Failed to acquire MSAL token: {exc}") from exc

    if not result or "access_token" not in result:
        message = result.get("error_description") if isinstance(result, dict) else "Unknown error"
        raise GraphClientError(f"MSAL did not return an access token: {message}")

    return result["access_token"]


def _download_drive_item(download_url: str) -> bytes:
    response = requests.get(download_url)
    try:
        response.raise_for_status()
    except requests.HTTPError as exc:
        raise GraphClientError(f"Failed to download file content: {exc}") from exc
    return response.content


def download_excel_from_path(access_token: str, site_id: str, file_path: str) -> bytes:
    """Download an Excel workbook from SharePoint by *file_path*."""

    if not site_id or not file_path:
        raise GraphClientError("SharePoint site identifier or file path is missing")

    headers = {"Authorization": f"Bearer {access_token}"}
    metadata_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_path}:/"

    metadata_response = requests.get(metadata_url, headers=headers)
    if metadata_response.status_code != 200:
        raise GraphClientError(
            f"Microsoft Graph returned HTTP {metadata_response.status_code} for metadata request"
        )

    download_url = metadata_response.json().get("@microsoft.graph.downloadUrl")
    if not download_url:
        raise GraphClientError("Microsoft Graph response did not include a download URL")

    return _download_drive_item(download_url)


def download_excel_from_file_id(access_token: str, site_id: str, file_id: str) -> bytes:
    """Download an Excel workbook from SharePoint using the file's *file_id*."""

    if not site_id or not file_id:
        raise GraphClientError("SharePoint site identifier or file id is missing")

    headers = {"Authorization": f"Bearer {access_token}"}
    metadata_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{file_id}"

    metadata_response = requests.get(metadata_url, headers=headers)
    if metadata_response.status_code != 200:
        raise GraphClientError(
            f"Microsoft Graph returned HTTP {metadata_response.status_code} for metadata request"
        )

    download_url = metadata_response.json().get("@microsoft.graph.downloadUrl")
    if not download_url:
        raise GraphClientError("Microsoft Graph response did not include a download URL")

    return _download_drive_item(download_url)
