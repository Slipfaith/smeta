from __future__ import annotations

from pathlib import Path

import resource_utils


def test_resource_path_uses_package_directory():
    relative = "templates/example.txt"
    expected = Path(resource_utils.__file__).resolve().parent / relative
    assert resource_utils.resource_path(relative) == expected


def test_resource_path_uses_meipass_when_available(monkeypatch, tmp_path):
    base = tmp_path / "bundle"
    base.mkdir()
    monkeypatch.setattr(resource_utils.sys, "_MEIPASS", str(base), raising=False)

    result = resource_utils.resource_path("foo/bar.txt")
    assert result == base / "foo/bar.txt"
