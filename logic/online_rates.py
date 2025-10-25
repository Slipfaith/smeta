"""Utilities for downloading and parsing rate tables from online sources."""

from __future__ import annotations

import logging
import os
from dataclasses import dataclass
from functools import lru_cache
from io import BytesIO
from typing import Dict, Iterable, Mapping, Optional, Tuple

import pandas as pd

from . import rates_importer
from .ms_graph_client import (
    GraphClientError,
    authenticate_with_msal,
    download_excel_from_file_id,
    download_excel_from_path,
)

from .env_loader import load_application_env

logger = logging.getLogger(__name__)
load_application_env()

RatesMap = rates_importer.RatesMap


@dataclass(frozen=True)
class RemoteSource:
    """Description of a remote rate workbook."""

    key: str
    label: str
    site_id: str
    parser: str
    file_path: Optional[str] = None
    file_id: Optional[str] = None
    sheet_name: Optional[str] = None
    skiprows: int = 0


SUPPORTED_TEP_COLUMNS: Mapping[str, Mapping[str, Tuple[str, str, str]]] = {
    "USD": {
        "R1": ("USD_Basic_R1", "USD_Complex_R1", "USD_Hourly_R1"),
        "R2": ("USD_Basic_R2", "USD_Complex_R2", "USD_Hourly_R2"),
    },
    "RUB": {
        "R1": ("RUB_Basic_R1", "RUB_Complex_R1", "RUB_Hourly_R1"),
    },
}


@lru_cache()
def _load_sources() -> Dict[str, RemoteSource]:
    sources: Dict[str, RemoteSource] = {}

    site_id_1 = os.getenv("SITE_ID_1")
    file_path_1 = os.getenv("FILE_PATH_1")
    if site_id_1 and file_path_1:
        sources["sharepoint_mlv"] = RemoteSource(
            key="sharepoint_mlv",
            label="SharePoint: MLV Rates",
            site_id=site_id_1,
            parser="mlv",
            file_path=file_path_1,
        )

    site_id_2 = os.getenv("SITE_ID_2")
    file_id_2 = os.getenv("FILE_ID_2")
    if site_id_2 and file_id_2:
        sources["sharepoint_tep"] = RemoteSource(
            key="sharepoint_tep",
            label="SharePoint: TEP (Source RU)",
            site_id=site_id_2,
            parser="tep",
            file_id=file_id_2,
            sheet_name="TEP (Source RU)",
            skiprows=3,
        )

    return sources


def available_sources() -> Mapping[str, RemoteSource]:
    """Return mapping of available remote rate sources configured in the environment."""

    return _load_sources()


def refresh_sources_cache() -> None:
    """Clear cached configuration for remote sources (useful for tests)."""

    _load_sources.cache_clear()


@lru_cache(maxsize=4)
def _download_source_bytes(source_key: str) -> bytes:
    sources = available_sources()
    source = sources.get(source_key)
    if not source:
        raise KeyError(f"Unknown remote source: {source_key}")

    client_id = os.getenv("CLIENT_ID")
    tenant_id = os.getenv("TENANT_ID")
    scope = os.getenv("SCOPE")
    if not client_id or not tenant_id or not scope:
        raise GraphClientError("Microsoft Graph credentials are not fully configured")

    scopes: Iterable[str] = [scope]
    token = authenticate_with_msal(client_id, tenant_id, scopes)

    if source.file_path:
        return download_excel_from_path(token, source.site_id, source.file_path)
    if source.file_id:
        return download_excel_from_file_id(token, source.site_id, source.file_id)
    raise GraphClientError("Remote source is missing both file path and file id")


def clear_download_cache() -> None:
    """Clear cached file downloads (useful when credentials change)."""

    _download_source_bytes.cache_clear()


def load_remote_rates(source_key: str, currency: str, rate_type: str) -> RatesMap:
    """Load rate table from the remote source *source_key*."""

    sources = available_sources()
    source = sources.get(source_key)
    if not source:
        raise KeyError(f"Unknown remote source: {source_key}")

    raw_bytes = _download_source_bytes(source_key)

    if source.parser == "mlv":
        return rates_importer.load_rates_from_excel(BytesIO(raw_bytes), currency, rate_type)
    if source.parser == "tep":
        return _load_tep_rates(BytesIO(raw_bytes), currency, rate_type, source)

    raise ValueError(f"Unsupported parser type: {source.parser}")


def _normalize_language(value: str) -> str:
    return rates_importer._normalize_language(value)


def _load_tep_rates(
    stream: BytesIO,
    currency: str,
    rate_type: str,
    source: RemoteSource,
) -> RatesMap:
    if not source.sheet_name:
        raise ValueError("TEP source requires a sheet name")

    df = pd.read_excel(
        stream,
        sheet_name=source.sheet_name,
        skiprows=source.skiprows,
        header=None,
    )

    rename_map = {}
    columns = list(df.columns)
    if len(columns) >= 11:
        rename_map[columns[0]] = "SourceLang"
        rename_map[columns[1]] = "TargetLang"
        rename_map[columns[2]] = "USD_Basic_R1"
        rename_map[columns[3]] = "USD_Complex_R1"
        rename_map[columns[4]] = "USD_Hourly_R1"
        rename_map[columns[5]] = "USD_Basic_R2"
        rename_map[columns[6]] = "USD_Complex_R2"
        rename_map[columns[7]] = "USD_Hourly_R2"
        rename_map[columns[8]] = "RUB_Basic_R1"
        rename_map[columns[9]] = "RUB_Complex_R1"
        rename_map[columns[10]] = "RUB_Hourly_R1"
    df = df.rename(columns=rename_map)

    if "SourceLang" not in df.columns or "TargetLang" not in df.columns:
        raise ValueError("TEP workbook does not contain SourceLang/TargetLang columns")

    df = df.dropna(subset=["SourceLang", "TargetLang"], how="any")

    currency = currency.upper()
    rate_type = rate_type.upper()

    try:
        basic_col, complex_col, hour_col = SUPPORTED_TEP_COLUMNS[currency][rate_type]
    except KeyError as exc:
        raise ValueError(
            f"TEP rates do not support {currency} {rate_type} combinations"
        ) from exc

    for col in [basic_col, complex_col, hour_col]:
        if col not in df.columns:
            raise ValueError(f"TEP workbook is missing required column: {col}")

    numeric_cols = [basic_col, complex_col, hour_col]
    df[numeric_cols] = df[numeric_cols].apply(pd.to_numeric, errors="coerce")

    rates: RatesMap = {}
    for _, row in df.iterrows():
        src = str(row["SourceLang"]).strip()
        tgt = str(row["TargetLang"]).strip()
        if not src or not tgt:
            continue

        basic = row.get(basic_col)
        complex_ = row.get(complex_col)
        hour = row.get(hour_col)

        if pd.isna(basic) and pd.isna(complex_) and pd.isna(hour):
            continue

        src_code = _normalize_language(src)
        tgt_code = _normalize_language(tgt)

        rates[(src_code, tgt_code)] = {
            "basic": float(basic) if not pd.isna(basic) else 0.0,
            "complex": float(complex_) if not pd.isna(complex_) else 0.0,
            "hour": float(hour) if not pd.isna(hour) else 0.0,
        }

    if not rates:
        raise ValueError("No usable rates found in the TEP workbook")

    return rates
