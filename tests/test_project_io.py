from __future__ import annotations

import json

from logic import project_io


def test_save_project_writes_json(tmp_path):
    data = {"name": "Example", "items": [1, 2, 3]}
    target = tmp_path / "project.json"

    assert project_io.save_project(data, target) is True

    loaded = json.loads(target.read_text(encoding="utf-8"))
    assert loaded == data


def test_save_project_returns_false_on_error(tmp_path):
    data = {"key": "value"}
    target_dir = tmp_path / "dir"
    target_dir.mkdir()

    assert project_io.save_project(data, target_dir) is False


def test_load_project_reads_json(tmp_path):
    payload = {"currency": "USD"}
    source = tmp_path / "project.json"
    source.write_text(json.dumps(payload), encoding="utf-8")

    assert project_io.load_project(source) == payload


def test_load_project_returns_none_when_missing(tmp_path):
    assert project_io.load_project(tmp_path / "missing.json") is None


def test_load_project_returns_none_on_invalid_json(tmp_path):
    source = tmp_path / "broken.json"
    source.write_text("not json", encoding="utf-8")

    assert project_io.load_project(source) is None
