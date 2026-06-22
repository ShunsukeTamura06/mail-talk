"""実行形態（通常実行 / PyInstaller exe）に応じたパス解決。

A端末ではPyInstallerで固めた単一exeとして動かす（CLAUDE.md §14）。frozen時は:
- 読み取り専用の同梱リソース(static)は `sys._MEIPASS` から読む。
- 書き込み先(data/logs)はexeと同じフォルダに置く（共有フォルダから持ち込んだ
  exeの隣にDB・ログが溜まる＝持ち帰り・診断がしやすい）。
"""

from __future__ import annotations

import sys
from pathlib import Path


def is_frozen() -> bool:
    """PyInstaller等で固められた実行形態かどうか。"""
    return bool(getattr(sys, "frozen", False))


def base_writable_dir() -> Path:
    """書き込み可能な基準ディレクトリを返す。

    Returns:
        frozen時はexeのあるフォルダ、通常実行時はリポジトリルート。
    """
    if is_frozen():
        return Path(sys.executable).resolve().parent
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
