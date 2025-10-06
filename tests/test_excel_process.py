import sys
import types

from logic import excel_process


def _reset_tracking():
    excel_process.excel_instances.clear()
    excel_process._tracked_excel_objects.clear()  # type: ignore[attr-defined]


def test_register_excel_instance_skips_on_non_windows(monkeypatch):
    _reset_tracking()
    excel = types.SimpleNamespace(Hwnd=123)

    monkeypatch.setattr(excel_process.sys, "platform", "linux")

    excel_process.register_excel_instance(excel)

    assert excel_process.excel_instances == set()


def test_register_and_unregister_excel_instance(monkeypatch):
    _reset_tracking()

    excel = types.SimpleNamespace(pid=321)

    monkeypatch.setattr(excel_process.sys, "platform", "win32")
    monkeypatch.setattr(excel_process, "_get_excel_pid", lambda obj: obj.pid)

    excel_process.register_excel_instance(excel)

    assert excel_process.excel_instances == {321}
    assert excel_process._tracked_excel_objects[321] is excel  # type: ignore[index]

    excel_process.unregister_excel_instance(excel)

    assert excel_process.excel_instances == set()
    assert excel_process._tracked_excel_objects == {}  # type: ignore[comparison-overlap]


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
        self.closed = True


class _FakeWorkbooks:
    def __init__(self, workbook: _FakeWorkbook):
        self.workbook = workbook
        self.opened_path = None
        self.Count = 1

    def Open(self, path: str):
        self.opened_path = path
        return self.workbook

    def __iter__(self):
        yield self.workbook

    def __call__(self, index: int):  # pragma: no cover - compatibility shim
        if index != 1:
            raise IndexError("Only one workbook in fake collection")
        return self.workbook


class _FakeExcel:
    def __init__(self, workbook: _FakeWorkbook, pid: int | None = None):
        self.DecimalSeparator = ","
        self.ThousandsSeparator = " "
        self.UseSystemSeparators = True
        self.Visible = True
        self.DisplayAlerts = True
        self.workbook = workbook
        self.Workbooks = _FakeWorkbooks(workbook)
        self.quit_called = False
        self.pid = pid

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
    _reset_tracking()
    workbook = _FakeWorkbook()
    excel = _FakeExcel(workbook)

    _install_fake_win32(monkeypatch, excel)

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


def test_apply_separators_handles_failure_and_cleans_up(monkeypatch, tmp_path):
    _reset_tracking()
    workbook = _FakeWorkbook(should_fail=True)
    excel = _FakeExcel(workbook)

    _install_fake_win32(monkeypatch, excel)

    xlsx_path = tmp_path / "example.xlsx"
    xlsx_path.write_text("dummy")

    assert excel_process.apply_separators(str(xlsx_path), "ru") is False

    assert workbook.closed is True
    assert excel.quit_called is True
    assert excel.UseSystemSeparators is True


def test_close_tracked_excel_instances_closes_tracked_only(monkeypatch):
    _reset_tracking()

    workbook1 = _FakeWorkbook()
    workbook2 = _FakeWorkbook()
    excel1 = _FakeExcel(workbook1, pid=101)
    excel2 = _FakeExcel(workbook2, pid=202)

    monkeypatch.setattr(excel_process.sys, "platform", "win32")
    monkeypatch.setattr(excel_process, "_get_excel_pid", lambda obj: obj.pid)

    excel_process.register_excel_instance(excel1)
    excel_process.register_excel_instance(excel2)

    class _FakePsutil:
        @staticmethod
        def pid_exists(pid):
            return pid == 101

    monkeypatch.setattr(excel_process, "psutil", _FakePsutil)

    excel_process.close_tracked_excel_instances()

    assert workbook1.closed is True
    assert excel1.quit_called is True
    assert workbook2.closed is False
    assert excel2.quit_called is False
    assert excel_process.excel_instances == set()
    assert excel_process._tracked_excel_objects == {}  # type: ignore[comparison-overlap]
