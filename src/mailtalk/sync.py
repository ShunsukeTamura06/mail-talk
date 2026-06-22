"""同期ワーカー＋ステートマシン（CLAUDE.md §6）。

`起動中 → Outlook接続中 → 同期中(◯/◯件) → 仕分け中 → 準備完了`、異常時は
`エラー(原因つき)`。新しい順に取得し、差分同期で2日目以降を高速化する。
バックグラウンドスレッドで動かし、UIは`status()`でいつでも進捗を取得できる。
"""

from __future__ import annotations

import threading
from dataclasses import asdict, dataclass, field
from datetime import datetime, timedelta

from .aggregate import build_conversations
from .config import get_config
from .db import Database
from .notify import log_debug, notify_user
from .source import OutlookSource, get_default_source
from .triage import classify_into

# バッチコミット間隔（毎回commitせずN件ごとに確定して書き込みを高速化）。
_COMMIT_EVERY = 200

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

            last = None if full else self.db.get_state("last_sync_time")
            is_cold = last is None  # 初回（or full）はキャッシュ無し

            cold_days = get_config().cold_window_days
            if is_cold and not full:
                # コールドスタート: まず直近ウィンドウだけ取り込んで「準備完了」に
                # 近づけ、古い分は裏で後追いする（§6 プログレッシブ）。
                window_start = datetime.now() - timedelta(days=cold_days)
                self._set(STATE_SYNCING, processed=0, total=None,
                          message=f"直近{cold_days}日のメールを読み込み中…")
                n1, newest = self._fetch_range(
                    since=window_start, before=None,
                    base_msg=f"直近{cold_days}日のメールを読み込み中…",
                )
                if newest is not None:
                    self.db.set_state("last_sync_time", newest.isoformat())
                count = self._classify_and_store()

                if not get_config().backfill_old:
                    # 設定でバックフィル無効 → 直近のみで完了（最速）。
                    self._set(STATE_READY, processed=n1, total=n1,
                              message=f"準備完了。{count}件の会話を仕分けしました（直近{cold_days}日）。")
                    notify_user("info", f"準備完了。直近{cold_days}日分（{count}件）を仕分けしました。")
                    return

                # 仮の準備完了（古い分は裏で読み込み中）。
                self._set(STATE_READY, processed=n1, total=n1,
                          message=f"準備完了（直近{cold_days}日 {count}件）。古いメールを読み込み中…")
                notify_user(
                    "info",
                    f"準備完了。まず直近{cold_days}日分（{count}件）を仕分けしました。古い分は裏で読み込みます。",
                )

                # 古い分のバックフィル（state は READY のまま継続）。
                n2, _ = self._fetch_range(
                    since=None, before=window_start,
                    base_msg="古いメールを読み込み中…", state=STATE_READY,
                )
                count = self._classify_and_store()
                self.db.set_state("backfill_done", "1")
                self._set(STATE_READY, processed=n1 + n2, total=n1 + n2,
                          message=f"準備完了。{count}件の会話を仕分けしました。")
                notify_user("info", f"全期間の読み込みが完了しました（{count}件の会話）。")
            else:
                # 差分同期（2回目以降）または全件再取得(full)。
                since = None if full else datetime.fromisoformat(last)
                if since is not None:
                    log_debug(f"差分同期: since={since.isoformat()}")
                self._set(STATE_SYNCING, processed=0, total=None,
                          message="メールを読み込み中…")
                processed, newest = self._fetch_range(
                    since=since, before=None, base_msg="メールを読み込み中…"
                )
                if newest is not None:
                    self.db.set_state("last_sync_time", newest.isoformat())
                count = self._classify_and_store()
                self.db.set_state("status", STATE_READY)
                self._set(STATE_READY, processed=processed, total=processed,
                          message=f"準備完了。{count}件の会話を仕分けしました。")
                notify_user(
                    "info",
                    f"準備完了。{count}件の会話を仕分けしました。",
                    detail=f"新規/更新メール {processed}件",
                )
        except Exception as exc:  # noqa: BLE001 - 失敗は状態として可視化する
            self._set(STATE_ERROR, error=str(exc), message="読み込みに失敗しました。")
            notify_user(
                "error",
                "読み込みに失敗しました。Outlookが起動しているか確認してください。",
                detail=repr(exc),
            )

    def _fetch_range(
        self,
        *,
        since: datetime | None,
        before: datetime | None,
        base_msg: str,
        state: str = STATE_SYNCING,
    ) -> tuple[int, datetime | None]:
        """指定範囲のメールを取得してDBへUPSERTし、進捗を更新する。

        Args:
            since: この時刻より後のみ。
            before: この時刻より前のみ（バックフィル用）。
            base_msg: 進捗メッセージの接頭。
            state: 進捗更新時に設定するステート（バックフィルは READY のまま）。

        Returns:
            (取得件数, 取得した中で最新の受信時刻)。
        """
        processed = 0
        newest: datetime | None = None
        for m in self.source.iter_messages(since=since, before=before):
            # commit=Falseで保留し、N件ごと＋最後にまとめてcommit（書き込み高速化）。
            self.db.upsert_email(m, commit=False)
            processed += 1
            if newest is None or m.received_time > newest:
                newest = m.received_time
            if processed % _COMMIT_EVERY == 0:
                self.db.commit()
            if processed % 50 == 0:
                self._set(state, processed=processed, message=f"{base_msg} {processed}件")
        self.db.commit()
        return processed, newest

    def _classify_and_store(self) -> int:
        """DB上の全メールから会話を再構築・仕分けして保存し、会話数を返す。"""
        convs = build_conversations(self.db.all_messages())
        for c in convs:
            classify_into(c)
            self.db.upsert_conversation(c, commit=False)
        self.db.commit()
        return len(convs)
