from __future__ import annotations

from io import BytesIO

import pandas as pd
import pytest
from openpyxl import Workbook

from logic import online_rates


@pytest.fixture(autouse=True)
def reset_online_rates(monkeypatch):
    """Ensure caches are cleared and environment is pristine between tests."""
    for key in [
        "SITE_ID_1",
        "FILE_PATH_1",
        "SITE_ID_2",
        "FILE_ID_2",
        "CLIENT_ID",
        "TENANT_ID",
        "SCOPE",
    ]:
        monkeypatch.delenv(key, raising=False)
    online_rates.refresh_sources_cache()
    online_rates.clear_download_cache()
    yield
    online_rates.refresh_sources_cache()
    online_rates.clear_download_cache()


def test_available_sources_empty_when_not_configured():
    assert online_rates.available_sources() == {}


def test_load_remote_mlv_rates(monkeypatch):
    monkeypatch.setenv("SITE_ID_1", "site")
    monkeypatch.setenv("FILE_PATH_1", "rates.xlsx")
    monkeypatch.setenv("CLIENT_ID", "client")
    monkeypatch.setenv("TENANT_ID", "tenant")
    monkeypatch.setenv("SCOPE", "scope/.default")

    online_rates.refresh_sources_cache()
    online_rates.clear_download_cache()

    wb = Workbook()
    ws = wb.active
    ws.title = "R1_USD"
    ws.append(["Source", "Target", "Basic", "Complex", "Hour"])
    ws.append(["English", "Russian", 1.23, 2.34, 3.45])
    ws.append(["German", "English", 4.56, 5.67, 6.78])

    buffer = BytesIO()
    wb.save(buffer)
    payload = buffer.getvalue()

    calls = []

    def fake_auth(client_id, tenant_id, scopes):
        calls.append("auth")
        return "token"

    def fake_download(_token, site_id, file_path):
        calls.append((site_id, file_path))
        return payload

    monkeypatch.setattr(online_rates, "authenticate_with_msal", fake_auth)
    monkeypatch.setattr(online_rates, "download_excel_from_path", fake_download)

    rates = online_rates.load_remote_rates("sharepoint_mlv", "USD", "R1")

    assert calls == ["auth", ("site", "rates.xlsx")]
    assert rates[("en", "ru")]["basic"] == pytest.approx(1.23)
    assert rates[("de", "en")]["hour"] == pytest.approx(6.78)


def test_load_remote_tep_rates(monkeypatch):
    monkeypatch.setenv("SITE_ID_2", "site2")
    monkeypatch.setenv("FILE_ID_2", "file2")
    monkeypatch.setenv("CLIENT_ID", "client")
    monkeypatch.setenv("TENANT_ID", "tenant")
    monkeypatch.setenv("SCOPE", "scope/.default")

    online_rates.refresh_sources_cache()
    online_rates.clear_download_cache()

    data = [
        ["English", "Russian", 1.111, 2.222, 3, 4.444, 5.555, 6, 7.7, 8.8, 9],
        ["English", "German", 0.9, 1.8, 2, 3.3, 4.4, 5, 6.6, 7.7, 8.8],
    ]
    df = pd.DataFrame(data)
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(
            writer,
            sheet_name="TEP (Source RU)",
            index=False,
            header=False,
            startrow=3,
        )
    payload = buffer.getvalue()

    monkeypatch.setattr(online_rates, "authenticate_with_msal", lambda *args, **kwargs: "token")
    monkeypatch.setattr(online_rates, "download_excel_from_file_id", lambda *_: payload)

    rates = online_rates.load_remote_rates("sharepoint_tep", "USD", "R1")

    assert rates[("en", "ru")]["basic"] == pytest.approx(1.111)
    assert rates[("en", "de")]["complex"] == pytest.approx(1.8)

    with pytest.raises(ValueError):
        online_rates.load_remote_rates("sharepoint_tep", "EUR", "R1")
