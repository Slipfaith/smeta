import sys
import types

from logic import excel_process


class _FakeProc:
    def __init__(self, name: str, pid: int, terminate_raises: bool = False):
        self.info = {"name": name, "pid": pid}
        self.pid = pid
        self.terminated = False
        self.killed = False
        self._terminate_raises = terminate_raises

    def terminate(self):
        if self._terminate_raises:
            raise RuntimeError("terminate failed")
        self.terminated = True

    def wait(self, timeout=None):
        del timeout
        return True

    def kill(self):
        self.killed = True


def test_close_excel_processes_skips_on_non_windows(monkeypatch):
    called = False

    class _Sentinel:
        @staticmethod
        def process_iter(_):
            nonlocal called
            called = True
            return []

    monkeypatch.setattr(excel_process.sys, "platform", "linux")
    monkeypatch.setattr(excel_process, "psutil", _Sentinel)
    excel_process._managed_excel_pids.clear()

    excel_process.close_excel_processes()

    assert called is False


def test_close_excel_processes_terminates_matching_processes(monkeypatch):
    target = _FakeProc("EXCEL.EXE", pid=123)
    other = _FakeProc("notepad.exe", pid=456)

    class _FakePsutil:
        @staticmethod
        def process_iter(_):
            return [target, other]

    monkeypatch.setattr(excel_process.sys, "platform", "win32")
    monkeypatch.setattr(excel_process, "psutil", _FakePsutil)
    excel_process._managed_excel_pids.clear()
    excel_process._managed_excel_pids.add(target.pid)

    excel_process.close_excel_processes()

    assert target.terminated is True
    assert other.terminated is False
    assert target.killed is False
    assert target.pid not in excel_process._managed_excel_pids


def test_close_excel_processes_kills_when_terminate_fails(monkeypatch):
    target = _FakeProc("excel.exe", pid=999, terminate_raises=True)

    class _FakePsutil:
        @staticmethod
        def process_iter(_):
            return [target]

    monkeypatch.setattr(excel_process.sys, "platform", "win32")
    monkeypatch.setattr(excel_process, "psutil", _FakePsutil)
    excel_process._managed_excel_pids.clear()
    excel_process._managed_excel_pids.add(target.pid)

    excel_process.close_excel_processes()

    assert target.killed is True
    assert target.pid not in excel_process._managed_excel_pids


class _FakeWorkbook:
    def __init__(self, should_fail: bool = False):
        self.saved = False
        self.closed = False
        self._should_fail = should_fail

    def Save(self):
        self.saved = True
        if self._should_fail:
            raise RuntimeError("cannot save")

    def Close(self, save_changes):
        del save_changes
        self.closed = True


class _FakeWorkbooks:
    def __init__(self, workbook: _FakeWorkbook):
        self.workbook = workbook
        self.opened_path = None

    def Open(self, path: str):
        self.opened_path = path
        return self.workbook


class _FakeExcel:
    def __init__(self, workbook: _FakeWorkbook):
        self.DecimalSeparator = ","
        self.ThousandsSeparator = " "
        self.UseSystemSeparators = True
        self.Visible = True
        self.DisplayAlerts = True
        self.workbook = workbook
        self.Workbooks = _FakeWorkbooks(workbook)
        self.quit_called = False

    def Quit(self):
        self.quit_called = True


def _install_fake_win32(monkeypatch, excel_instance):
    fake_client = types.SimpleNamespace(Dispatch=lambda name: excel_instance)
    fake_module = types.SimpleNamespace(client=fake_client)
    monkeypatch.setitem(sys.modules, "win32com", fake_module)
    monkeypatch.setitem(sys.modules, "win32com.client", fake_client)


def test_temporary_separators_sets_and_restores():
    excel = types.SimpleNamespace(
        DecimalSeparator=",",
        ThousandsSeparator=" ",
        UseSystemSeparators=True,
    )

    with excel_process.temporary_separators(excel, "en-GB"):
        assert excel.DecimalSeparator == "."
        assert excel.ThousandsSeparator == ","
        assert excel.UseSystemSeparators is False

    assert excel.DecimalSeparator == ","
    assert excel.ThousandsSeparator == " "
    assert excel.UseSystemSeparators is True


def test_apply_separators_success(monkeypatch, tmp_path):
    workbook = _FakeWorkbook()
    excel = _FakeExcel(workbook)

    _install_fake_win32(monkeypatch, excel)
    monkeypatch.setattr(excel_process, "_get_excel_pid", lambda _: 321)
    excel_process._managed_excel_pids.clear()

    xlsx_path = tmp_path / "example.xlsx"
    xlsx_path.write_text("dummy")

    assert excel_process.apply_separators(str(xlsx_path), "en") is True

    assert excel.Workbooks.opened_path == str(xlsx_path)
    assert workbook.saved is True
    assert workbook.closed is True
    assert excel.quit_called is True
    assert excel.DecimalSeparator == ","
    assert excel.ThousandsSeparator == " "
    assert excel.UseSystemSeparators is True
    assert excel_process._managed_excel_pids == set()


def test_apply_separators_handles_failure_and_cleans_up(monkeypatch, tmp_path):
    workbook = _FakeWorkbook(should_fail=True)
    excel = _FakeExcel(workbook)

    _install_fake_win32(monkeypatch, excel)
    monkeypatch.setattr(excel_process, "_get_excel_pid", lambda _: 654)
    excel_process._managed_excel_pids.clear()

    xlsx_path = tmp_path / "example.xlsx"
    xlsx_path.write_text("dummy")

    assert excel_process.apply_separators(str(xlsx_path), "ru") is False

    assert workbook.closed is True
    assert excel.quit_called is True
    assert excel.UseSystemSeparators is True
    assert excel_process._managed_excel_pids == set()
