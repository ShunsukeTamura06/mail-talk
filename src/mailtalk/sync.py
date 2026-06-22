"""同期ワーカー＋ステートマシン（CLAUDE.md §6）。

`起動中 → Outlook接続中 → 同期中(◯/◯件) → 仕分け中 → 準備完了`、異常時は
`エラー(原因つき)`。新しい順に取得し、差分同期で2日目以降を高速化する。
バックグラウンドスレッドで動かし、UIは`status()`でいつでも進捗を取得できる。
"""

from __future__ import annotations

import threading
from dataclasses import asdict, dataclass, field
from datetime import datetime

from .aggregate import build_conversations
from .db import Database
from .notify import log_debug, notify_user
from .source import OutlookSource, get_default_source
from .triage import classify_into

# ステート定数。
STATE_STARTING = "starting"
STATE_CONNECTING = "connecting"
STATE_SYNCING = "syncing"
STATE_TRIAGING = "triaging"
STATE_READY = "ready"
STATE_ERROR = "error"

_STATE_LABELS = {
    STATE_STARTING: "起動中",
    STATE_CONNECTING: "Outlook接続中",
    STATE_SYNCING: "同期中",
    STATE_TRIAGING: "仕分け中",
    STATE_READY: "準備完了",
    STATE_ERROR: "エラー",
}


@dataclass
class SyncStatus:
    """同期の現在状態。UI/エンドポイントが参照する。"""

    state: str = STATE_STARTING
    label: str = _STATE_LABELS[STATE_STARTING]
    processed: int = 0
    total: int | None = None
    message: str = ""
    error: str | None = None
    updated_at: str = field(default_factory=lambda: datetime.now().isoformat())


class SyncManager:
    """同期の進捗を保持し、バックグラウンド同期を駆動する。"""

    def __init__(
        self, db: Database, source: OutlookSource | None = None
    ) -> None:
        self.db = db
        self.source = source or get_default_source()
        self._status = SyncStatus()
        self._lock = threading.Lock()
        self._thread: threading.Thread | None = None

    # -- 状態 ---------------------------------------------------------------

    def status(self) -> dict:
        """現在の同期状態をdictで返す。"""
        with self._lock:
            return asdict(self._status)

    def _set(
        self,
        state: str,
        *,
        processed: int | None = None,
        total: int | None = None,
        message: str = "",
        error: str | None = None,
    ) -> None:
        with self._lock:
            self._status.state = state
            self._status.label = _STATE_LABELS.get(state, state)
            if processed is not None:
                self._status.processed = processed
            if total is not None:
                self._status.total = total
            if message:
                self._status.message = message
            self._status.error = error
            self._status.updated_at = datetime.now().isoformat()

    # -- 実行 ---------------------------------------------------------------

    def start_background(self, full: bool = False) -> None:
        """同期をバックグラウンドスレッドで開始する。

        Args:
            full: Trueで全件再取得。Falseで差分同期（既定）。
        """
        if self._thread and self._thread.is_alive():
            log_debug("同期は既に実行中。新規起動をスキップ。")
            return
        self._thread = threading.Thread(
            target=self.run, kwargs={"full": full}, daemon=True
        )
        self._thread.start()

    def run(self, full: bool = False) -> None:
        """同期を同期的に実行する（スレッド本体）。

        Args:
            full: Trueで全件再取得。Falseで差分同期。
        """
        try:
            self._set(STATE_STARTING, message="準備を開始します。")

            # 1. 接続（Outlook起動待ちを含む）。
            self._set(STATE_CONNECTING, message="Outlookに接続しています。")
            if hasattr(self.source, "connect"):
                self.source.connect()

            # 2. 差分同期の起点を決める。
            since = None
            if not full:
                last = self.db.get_state("last_sync_time")
                if last:
                    since = datetime.fromisoformat(last)
                    log_debug(f"差分同期: since={since.isoformat()}")

            # 3. メール取得（新しい順）＋UPSERT。進捗を更新。
            self._set(STATE_SYNCING, processed=0, total=None, message="メールを読み込み中…")
            processed = 0
            newest_seen: datetime | None = None
            for m in self.source.iter_messages(since=since):
                self.db.upsert_email(m)
                processed += 1
                if newest_seen is None or m.received_time > newest_seen:
                    newest_seen = m.received_time
                if processed % 50 == 0:
                    self._set(STATE_SYNCING, processed=processed,
                              message=f"メールを読み込み中… {processed}件")

            # 4. 仕分け（全メールから会話を再構築）。
            self._set(STATE_TRIAGING, processed=processed, message="会話を仕分け中…")
            convs = build_conversations(self.db.all_messages())
            for c in convs:
                classify_into(c)
                self.db.upsert_conversation(c)

            # 5. 差分同期の起点を保存。
            if newest_seen is not None:
                self.db.set_state("last_sync_time", newest_seen.isoformat())
            self.db.set_state("status", STATE_READY)

            self._set(
                STATE_READY,
                processed=processed,
                total=processed,
                message=f"準備完了。{len(convs)}件の会話を仕分けしました。",
            )
            notify_user(
                "info",
                f"準備完了。{len(convs)}件の会話を仕分けしました。",
                detail=f"新規/更新メール {processed}件",
            )
        except Exception as exc:  # noqa: BLE001 - 失敗は状態として可視化する
            self._set(STATE_ERROR, error=str(exc), message="読み込みに失敗しました。")
            notify_user(
                "error",
                "読み込みに失敗しました。Outlookが起動しているか確認してください。",
                detail=repr(exc),
            )
