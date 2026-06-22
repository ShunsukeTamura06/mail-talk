"""外部設定ファイル(config.json)の読み込みテスト。"""

from __future__ import annotations

import json

from mailtalk.config import Config, ensure_config_file, load_config


def test_defaults_when_no_file(tmp_path):
    cfg = load_config(tmp_path / "nope.json")
    assert cfg == Config()
    assert cfg.cold_window_days == 90
    assert cfg.backfill_old is True


def test_load_overrides_from_file(tmp_path):
    p = tmp_path / "config.json"
    p.write_text(
        json.dumps({"cold_window_days": 30, "backfill_old": False, "port": 9000}),
        encoding="utf-8",
    )
    cfg = load_config(p)
    assert cfg.cold_window_days == 30
    assert cfg.backfill_old is False
    assert cfg.port == 9000
    assert cfg.host == "127.0.0.1"  # 未指定キーは既定


def test_unknown_keys_ignored_and_bad_types_fallback(tmp_path):
    p = tmp_path / "config.json"
    p.write_text(
        json.dumps({"cold_window_days": "abc", "unknown": 1}),
        encoding="utf-8",
    )
    cfg = load_config(p)
    assert cfg.cold_window_days == 90  # 不正な型は既定にフォールバック


def test_ensure_creates_file_without_overwriting(tmp_path):
    p = tmp_path / "config.json"
    ensure_config_file(p)
    assert p.exists()
    data = json.loads(p.read_text(encoding="utf-8"))
    assert data["cold_window_days"] == 90

    # 既存は上書きしない。
    p.write_text(json.dumps({"cold_window_days": 7}), encoding="utf-8")
    ensure_config_file(p)
    assert json.loads(p.read_text(encoding="utf-8"))["cold_window_days"] == 7
