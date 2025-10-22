import sys
from types import ModuleType, SimpleNamespace

import pytest

from logic import outlook_com_cache


@pytest.fixture(autouse=True)
def restore_sys_platform(monkeypatch):
    monkeypatch.setattr(outlook_com_cache, "sys", SimpleNamespace(platform="win32"))


def test_rebuild_outlook_com_cache_noop_on_non_windows(monkeypatch):
    monkeypatch.setattr(outlook_com_cache.sys, "platform", "linux")

    outlook_com_cache.rebuild_outlook_com_cache()


def test_rebuild_outlook_com_cache_removes_gen_py_on_failure(monkeypatch, tmp_path):
    gen_py = tmp_path / "gen_py"
    gen_py.mkdir()

    ensure_calls = []

    def ensure_dispatch(prog_id):
        ensure_calls.append(prog_id)
        if len(ensure_calls) == 1:
            raise RuntimeError("broken cache")
        return SimpleNamespace()

    fake_client = ModuleType("win32com.client")
    fake_client.gencache = SimpleNamespace(EnsureDispatch=ensure_dispatch)

    fake_win32com = ModuleType("win32com")
    fake_win32com.__gen_path__ = str(gen_py)
    fake_win32com.client = fake_client

    pythoncom_calls = []

    fake_pythoncom = ModuleType("pythoncom")

    def co_init():
        pythoncom_calls.append("init")

    def co_uninit():
        pythoncom_calls.append("uninit")

    fake_pythoncom.CoInitialize = co_init
    fake_pythoncom.CoUninitialize = co_uninit

    monkeypatch.setitem(sys.modules, "win32com", fake_win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", fake_client)
    monkeypatch.setitem(sys.modules, "pythoncom", fake_pythoncom)

    outlook_com_cache.rebuild_outlook_com_cache()

    assert ensure_calls == ["Outlook.Application", "Outlook.Application"]
    assert not gen_py.exists()
    assert pythoncom_calls in (["init", "uninit"], ["init"])  # CoUninitialize optional if init failed
