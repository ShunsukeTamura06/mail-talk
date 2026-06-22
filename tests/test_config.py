"""外部設定ファイル(config.json)の読み込みテスト。"""

from __future__ import annotations

import json

from mailtalk.config import Config, enforce_loopback, ensure_config_file, load_config


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


def test_enforce_loopback_blocks_external_hosts():
    # ループバックはそのまま通す。
    assert enforce_loopback("127.0.0.1") == "127.0.0.1"
    assert enforce_loopback("127.0.0.5") == "127.0.0.5"
    assert enforce_loopback("localhost") == "localhost"
    assert enforce_loopback("::1") == "::1"
    # 外部公開につながる指定は 127.0.0.1 へ強制。
    assert enforce_loopback("0.0.0.0") == "127.0.0.1"
    assert enforce_loopback("192.168.1.10") == "127.0.0.1"
    assert enforce_loopback("") == "127.0.0.1"
    assert enforce_loopback("example.com") == "127.0.0.1"


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
