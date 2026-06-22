"""実行形態（通常実行 / PyInstaller exe）に応じたパス解決。

M端末ではPyInstallerで固めた単一exeとして動かす（CLAUDE.md §14）。frozen時は:
- 読み取り専用の同梱リソース(static)は `sys._MEIPASS` から読む。
- 書き込み先(data/logs)はexeと同じフォルダに置く（共有フォルダから持ち込んだ
  exeの隣にDB・ログが溜まる＝持ち帰り・診断がしやすい）。
"""

from __future__ import annotations

import os
import sys
from pathlib import Path


def is_frozen() -> bool:
    """PyInstaller等で固められた実行形態かどうか。"""
    return bool(getattr(sys, "frozen", False))


def _is_writable(d: Path) -> bool:
    """ディレクトリに書き込めるか実際に試す。"""
    try:
        d.mkdir(parents=True, exist_ok=True)
        probe = d / ".write_test"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
        return True
    except Exception:  # noqa: BLE001
        return False


def _local_appdata_dir() -> Path:
    """%LOCALAPPDATA%/MailTalk（無ければホーム配下）を返す。"""
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
    root = Path(base) if base else Path.home()
    return root / "MailTalk"


def base_writable_dir() -> Path:
    """書き込み可能な基準ディレクトリを返す。

    frozen時はまずexeの隣を試し、書き込めなければ（読み取り専用フォルダや
    共有フォルダ上での起動）%LOCALAPPDATA%/MailTalk へフォールバックする。

    Returns:
        書き込み可能な基準ディレクトリ。通常実行時はリポジトリルート。
    """
    if is_frozen():
        beside_exe = Path(sys.executable).resolve().parent
        if _is_writable(beside_exe):
            return beside_exe
        return _local_appdata_dir()
    return Path(__file__).resolve().parents[2]


def data_dir() -> Path:
    """SQLite等の永続データ置き場（作成はしない）。"""
    return base_writable_dir() / "data"


def logs_dir() -> Path:
    """ログ置き場（作成はしない）。"""
    return base_writable_dir() / "logs"


def static_dir() -> Path:
    """同梱の静的アセット(static)ディレクトリ。

    Returns:
        frozen時は `sys._MEIPASS/static`、通常時はパッケージ内 static。
    """
    if is_frozen():
        return Path(getattr(sys, "_MEIPASS", base_writable_dir())) / "static"
    return Path(__file__).resolve().parent / "static"
