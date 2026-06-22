"""二層メッセージ／ログ基盤（CLAUDE.md §7）。

1か所の呼び出しで行き先を2つに分ける:
- ユーザー向け: 平易な日本語の状態・警告。UIが読めるよう保持する。
- 開発者向け: 詳細ログをローテーション付きでファイル(logs/app.log)へ。

`notify_user(level, message, detail=...)` を呼ぶと両方へ適切な粒度で流れる。
シークレット（トークン/APIキーの値）はここに渡さない設計とする。
"""

from __future__ import annotations

import logging
from collections import deque
from dataclasses import asdict, dataclass
from datetime import datetime
from logging.handlers import RotatingFileHandler

from .paths import logs_dir

# ログ出力先。通常実行はリポジトリ直下、exe実行時はexeの隣の logs/。
_LOG_DIR = logs_dir()
_LOG_FILE = _LOG_DIR / "app.log"

_USER_LEVELS = ("info", "warn", "error")

# 直近のユーザー向けメッセージを保持（UI/状態エンドポイントが参照）。
_user_messages: deque[UserMessage] = deque(maxlen=200)

_logger: logging.Logger | None = None


@dataclass
class UserMessage:
    """UI表示用のユーザー向けメッセージ1件。"""

    level: str
    message: str
    at: str  # ISO8601


def get_logger() -> logging.Logger:
    """開発者向けローテーションログのロガーを返す（初回に初期化）。

    Returns:
        `logs/app.log` へ出力するロガー。
    """
    global _logger
    if _logger is not None:
        return _logger

    _LOG_DIR.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("mailtalk")
    logger.setLevel(logging.DEBUG)
    logger.propagate = False

    if not logger.handlers:
        handler = RotatingFileHandler(
            _LOG_FILE, maxBytes=2_000_000, backupCount=5, encoding="utf-8"
        )
        fmt = logging.Formatter(
            "%(asctime)s [%(levelname)s] %(message)s",
            datefmt="%Y-%m-%dT%H:%M:%S",
        )
        handler.setFormatter(fmt)
        logger.addHandler(handler)

        # 開発機では標準エラーにも出すと進捗が見やすい。
        stream = logging.StreamHandler()
        stream.setLevel(logging.INFO)
        stream.setFormatter(fmt)
        logger.addHandler(stream)

    _logger = logger
    return logger


def notify_user(level: str, message: str, *, detail: str | None = None) -> None:
    """ユーザー向け通知と開発者向けログへ同時に流す。

    Args:
        level: "info" / "warn" / "error" のいずれか。
        message: ユーザーに見せる平易な日本語の一文。
        detail: 開発者向けの技術的詳細（COMエラーコード、所要時間など）。
            UIには出さずログにのみ残す。シークレットは渡さないこと。
    """
    if level not in _USER_LEVELS:
        level = "info"

    _user_messages.append(
        UserMessage(level=level, message=message, at=datetime.now().isoformat())
    )

    logger = get_logger()
    log_line = message if detail is None else f"{message} | {detail}"
    if level == "error":
        logger.error(log_line)
    elif level == "warn":
        logger.warning(log_line)
    else:
        logger.info(log_line)


def log_debug(message: str) -> None:
    """開発者向けにのみ詳細ログを残す（ユーザーには出さない）。

    Args:
        message: 技術的な詳細メッセージ。
    """
    get_logger().debug(message)


def get_user_messages(limit: int = 50) -> list[dict]:
    """直近のユーザー向けメッセージを新しい順で返す。

    Args:
        limit: 返す最大件数。

    Returns:
        `UserMessage` をdict化したリスト（新しい順）。
    """
    items = list(_user_messages)[-limit:]
    items.reverse()
    return [asdict(m) for m in items]
