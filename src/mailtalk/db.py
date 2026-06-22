"""SQLite キャッシュ層（CLAUDE.md §8）。

メール内容を含むため、生成されるDBファイルはコミットしない（.gitignore済み）。
スキーマは §8 のたたき台に準拠。`emails`（生メール）/`conversations`（集約＋
レーン）/`contacts`/`sync_state`（差分同期の状態）。
"""

from __future__ import annotations

import json
import sqlite3
from datetime import datetime
from pathlib import Path

from .models import Conversation, Message

# 既定のDBパス。リポジトリルート直下の data/。
DEFAULT_DB_PATH = Path(__file__).resolve().parents[2] / "data" / "mailtalk.db"

_SCHEMA = """
CREATE TABLE IF NOT EXISTS emails (
  entry_id         TEXT PRIMARY KEY,
  store_id         TEXT,
  conversation_id  TEXT,
  subject          TEXT,
  subject_norm     TEXT,
  sender_email     TEXT,
  sender_name      TEXT,
  to_json          TEXT,
  cc_json          TEXT,
  is_to_me         INTEGER,
  is_cc_me         INTEGER,
  is_from_me       INTEGER,
  body_preview     TEXT,
  body_html        TEXT,
  received_time    TEXT,
  unread           INTEGER,
  importance       INTEGER,
  folder           TEXT,
  synced_at        TEXT
);
CREATE INDEX IF NOT EXISTS idx_emails_conv ON emails(conversation_id);
CREATE INDEX IF NOT EXISTS idx_emails_received ON emails(received_time);

CREATE TABLE IF NOT EXISTS conversations (
  conversation_id   TEXT PRIMARY KEY,
  subject_norm      TEXT,
  participants_json TEXT,
  participant_count INTEGER,
  last_received     TEXT,
  last_sender_email TEXT,
  last_from_me      INTEGER,
  any_unread        INTEGER,
  i_am_to           INTEGER,
  i_am_cc_only      INTEGER,
  velocity_recent   INTEGER,
  lane              TEXT,
  lane_reason       TEXT,
  updated_at        TEXT
);
CREATE INDEX IF NOT EXISTS idx_conv_lane ON conversations(lane);
CREATE INDEX IF NOT EXISTS idx_conv_last ON conversations(last_received);

CREATE TABLE IF NOT EXISTS contacts (
  email        TEXT PRIMARY KEY,
  display_name TEXT
);

CREATE TABLE IF NOT EXISTS sync_state (
  key   TEXT PRIMARY KEY,
  value TEXT
);
"""


class Database:
    """SQLite接続のラッパ。スキーマ管理・読み書きを担う。"""

    def __init__(self, path: str | Path = DEFAULT_DB_PATH) -> None:
        self.path = Path(path)
        if str(self.path) != ":memory:":
            self.path.parent.mkdir(parents=True, exist_ok=True)
        # 同期はバックグラウンドスレッドで動くため、スレッドをまたいで接続を共有する。
        # SQLiteの既定はserializedモードのため、C層のミューテックスで直列化される。
        self.conn = sqlite3.connect(str(self.path), check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.conn.executescript(_SCHEMA)
        self.conn.commit()

    def close(self) -> None:
        """接続を閉じる。"""
        self.conn.close()

    # -- emails -------------------------------------------------------------

    def upsert_email(self, m: Message) -> None:
        """メール1件をUPSERTする。

        Args:
            m: 保存するメール。
        """
        self.conn.execute(
            """
            INSERT INTO emails (
              entry_id, store_id, conversation_id, subject, subject_norm,
              sender_email, sender_name, to_json, cc_json,
              is_to_me, is_cc_me, is_from_me, body_preview, body_html,
              received_time, unread, importance, folder, synced_at
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(entry_id) DO UPDATE SET
              conversation_id=excluded.conversation_id,
              subject=excluded.subject,
              subject_norm=excluded.subject_norm,
              sender_email=excluded.sender_email,
              sender_name=excluded.sender_name,
              to_json=excluded.to_json,
              cc_json=excluded.cc_json,
              is_to_me=excluded.is_to_me,
              is_cc_me=excluded.is_cc_me,
              is_from_me=excluded.is_from_me,
              body_preview=excluded.body_preview,
              body_html=excluded.body_html,
              received_time=excluded.received_time,
              unread=excluded.unread,
              importance=excluded.importance,
              folder=excluded.folder,
              synced_at=excluded.synced_at
            """,
            (
                m.entry_id, m.store_id, m.conversation_id, m.subject, m.subject_norm,
                m.sender_email, m.sender_name,
                json.dumps(m.to_list), json.dumps(m.cc_list),
                int(m.is_to_me), int(m.is_cc_me), int(m.is_from_me),
                m.body_preview, m.body_html,
                m.received_time.isoformat(), int(m.unread), m.importance, m.folder,
                datetime.now().isoformat(),
            ),
        )
        self.conn.commit()

    def messages_for_conversation(self, conversation_id: str) -> list[Message]:
        """会話に属するメールを時系列（古い順）で返す。

        Args:
            conversation_id: 会話ID。

        Returns:
            時系列順の `Message` リスト。
        """
        rows = self.conn.execute(
            "SELECT * FROM emails WHERE conversation_id = ? ORDER BY received_time ASC",
            (conversation_id,),
        ).fetchall()
        return [_row_to_message(r) for r in rows]

    def all_messages(self) -> list[Message]:
        """保存済み全メールを返す（再集約用）。"""
        rows = self.conn.execute("SELECT * FROM emails").fetchall()
        return [_row_to_message(r) for r in rows]

    # -- conversations ------------------------------------------------------

    def upsert_conversation(self, c: Conversation) -> None:
        """会話（集約＋レーン）をUPSERTする。

        Args:
            c: 保存する会話（`lane`/`lane_reason`確定済み）。
        """
        self.conn.execute(
            """
            INSERT INTO conversations (
              conversation_id, subject_norm, participants_json, participant_count,
              last_received, last_sender_email, last_from_me, any_unread,
              i_am_to, i_am_cc_only, velocity_recent, lane, lane_reason, updated_at
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(conversation_id) DO UPDATE SET
              subject_norm=excluded.subject_norm,
              participants_json=excluded.participants_json,
              participant_count=excluded.participant_count,
              last_received=excluded.last_received,
              last_sender_email=excluded.last_sender_email,
              last_from_me=excluded.last_from_me,
              any_unread=excluded.any_unread,
              i_am_to=excluded.i_am_to,
              i_am_cc_only=excluded.i_am_cc_only,
              velocity_recent=excluded.velocity_recent,
              lane=excluded.lane,
              lane_reason=excluded.lane_reason,
              updated_at=excluded.updated_at
            """,
            (
                c.conversation_id, c.subject_norm,
                json.dumps(c.participants), c.participant_count,
                c.last_received.isoformat() if c.last_received else None,
                c.last_sender_email, int(c.last_from_me), int(c.any_unread),
                int(c.i_am_to), int(c.i_am_cc_only), c.velocity_recent,
                c.lane, c.lane_reason, datetime.now().isoformat(),
            ),
        )
        self.conn.commit()

    def conversations(self, lane: str | None = None) -> list[dict]:
        """会話一覧をUI向けdictで返す（最終受信が新しい順）。

        Args:
            lane: 指定するとそのレーンのみ。Noneで全件。

        Returns:
            会話のdictリスト。
        """
        if lane:
            rows = self.conn.execute(
                "SELECT * FROM conversations WHERE lane = ? ORDER BY last_received DESC",
                (lane,),
            ).fetchall()
        else:
            rows = self.conn.execute(
                "SELECT * FROM conversations ORDER BY last_received DESC"
            ).fetchall()
        return [_row_to_conv_dict(r) for r in rows]

    # -- sync_state ---------------------------------------------------------

    def set_state(self, key: str, value: str) -> None:
        """同期状態のキー値を保存する。"""
        self.conn.execute(
            "INSERT INTO sync_state (key, value) VALUES (?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (key, value),
        )
        self.conn.commit()

    def get_state(self, key: str, default: str | None = None) -> str | None:
        """同期状態のキー値を取得する。"""
        row = self.conn.execute(
            "SELECT value FROM sync_state WHERE key = ?", (key,)
        ).fetchone()
        return row["value"] if row else default


def _row_to_message(r: sqlite3.Row) -> Message:
    """emails行を`Message`へ復元する。"""
    m = Message(
        entry_id=r["entry_id"],
        store_id=r["store_id"] or "",
        conversation_id=r["conversation_id"] or "",
        subject=r["subject"] or "",
        sender_email=r["sender_email"] or "",
        sender_name=r["sender_name"] or "",
        to_list=json.loads(r["to_json"] or "[]"),
        cc_list=json.loads(r["cc_json"] or "[]"),
        received_time=datetime.fromisoformat(r["received_time"]),
        body_preview=r["body_preview"] or "",
        body_html=r["body_html"] or "",
        unread=bool(r["unread"]),
        importance=r["importance"] or 1,
        folder=r["folder"] or "",
    )
    m.is_to_me = bool(r["is_to_me"])
    m.is_cc_me = bool(r["is_cc_me"])
    m.is_from_me = bool(r["is_from_me"])
    return m


def _row_to_conv_dict(r: sqlite3.Row) -> dict:
    """conversations行をUI向けdictへ変換する。"""
    return {
        "conversation_id": r["conversation_id"],
        "subject_norm": r["subject_norm"],
        "participants": json.loads(r["participants_json"] or "[]"),
        "participant_count": r["participant_count"],
        "last_received": r["last_received"],
        "last_sender_email": r["last_sender_email"],
        "last_from_me": bool(r["last_from_me"]),
        "any_unread": bool(r["any_unread"]),
        "i_am_to": bool(r["i_am_to"]),
        "i_am_cc_only": bool(r["i_am_cc_only"]),
        "velocity_recent": r["velocity_recent"],
        "lane": r["lane"],
        "lane_reason": r["lane_reason"],
    }
