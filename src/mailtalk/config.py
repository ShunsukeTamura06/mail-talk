"""外部設定ファイル(config.json)の読み込み（CLAUDE.md §14）。

exe（または開発機ではリポジトリ）の隣に置く `config.json` で挙動を調整する。
M端末ではファイルを編集して再起動するだけで設定を変えられる。ファイルが無ければ
既定値で動き、`ensure_config_file()` を呼ぶと既定の設定ファイルを書き出す
（起動時に1回だけ生成。ユーザーの編集は上書きしない）。
"""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass

from .notify import log_debug
from .paths import base_writable_dir

CONFIG_FILENAME = "config.json"


@dataclass
class Config:
    """アプリの調整可能な設定。既定値は実運用の出発点。"""

    cold_window_days: int = 90  # 初回にまず取り込む直近日数
    backfill_old: bool = True  # 直近より古いメールを裏で後追い取得するか
    active_threshold: int = 6  # 🟠活発と判定する直近件数のしきい値
    # Toが解決不能(配布リスト/EX変換失敗)のとき🔴寄りに倒すか。trueにすると
    # 再現率は上がるが、解決失敗が多い環境では🔴が暴発するため既定はfalse。
    red_on_unresolved_to: bool = False
    host: str = "127.0.0.1"  # 待受ホスト（外部公開しないため既定固定）
    port: int = 8765
    open_browser: bool = True  # 起動時に既定ブラウザを開くか


_config: Config | None = None


def _config_path():
    return base_writable_dir() / CONFIG_FILENAME


def _coerce(raw: dict) -> Config:
    """dictを既定値とマージしつつ型を整えて Config にする（未知キーは無視）。"""
    d = Config()
    for key in (
        "cold_window_days",
        "backfill_old",
        "active_threshold",
        "red_on_unresolved_to",
        "host",
        "port",
        "open_browser",
    ):
        if key in raw and raw[key] is not None:
            cur = getattr(d, key)
            try:
                if isinstance(cur, bool):
                    setattr(d, key, bool(raw[key]))
                elif isinstance(cur, int):
                    setattr(d, key, int(raw[key]))
                else:
                    setattr(d, key, str(raw[key]))
            except (TypeError, ValueError):
                log_debug(f"config: {key} の値が不正のため既定を使用: {raw[key]!r}")
    return d


def load_config(path=None) -> Config:
    """設定ファイルを読み込む（無ければ既定値）。ファイルは書き出さない。

    Args:
        path: 読み込むパス。省略時は base_writable_dir()/config.json。

    Returns:
        読み込んだ（または既定の）Config。
    """
    p = path or _config_path()
    try:
        if p.exists():
            raw = json.loads(p.read_text(encoding="utf-8"))
            return _coerce(raw)
    except Exception as exc:  # noqa: BLE001 - 壊れた設定でも既定で起動継続
        log_debug(f"config読み込み失敗（既定値を使用）: {exc!r}")
    return Config()


def ensure_config_file(path=None) -> None:
    """設定ファイルが無ければ既定値で生成する（既存は上書きしない）。"""
    p = path or _config_path()
    try:
        if not p.exists():
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_text(
                json.dumps(asdict(Config()), ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            log_debug(f"既定の設定ファイルを生成: {p}")
    except Exception as exc:  # noqa: BLE001
        log_debug(f"config生成失敗: {exc!r}")


# ループバック（同一マシン内のみ到達可能）と認めるホスト。これ以外は外部公開に
# つながりうるため、設定で指定されても 127.0.0.1 へ強制する（職場ポリシー §1）。
_LOOPBACK_NAMES = {"localhost", "::1"}


def enforce_loopback(host: str) -> str:
    """外部公開を防ぐため、ループバック以外のホスト指定を 127.0.0.1 へ丸める。

    Args:
        host: 設定されたホスト。

    Returns:
        ループバックなら元の値、そうでなければ "127.0.0.1"。
    """
    h = (host or "").strip().lower()
    if h in _LOOPBACK_NAMES or h.startswith("127."):
        return host
    return "127.0.0.1"


def get_config() -> Config:
    """現在の設定を返す（初回に読み込み、以後キャッシュ）。"""
    global _config
    if _config is None:
        _config = load_config()
    return _config


def set_config(cfg: Config) -> None:
    """設定を差し替える（主にテスト用）。"""
    global _config
    _config = cfg


def reset_config() -> None:
    """キャッシュを破棄して次回再読み込みさせる。"""
    global _config
    _config = None
