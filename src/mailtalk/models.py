"""ドメインモデル。COM非依存の純粋なデータ構造。

このモジュールは取得層（win32com）から完全に独立しており、macOS開発機でも
Fakeデータで全ロジックを検証できる。`outlook_client`（実機）も`fake_outlook`
（開発機）も、最終的にここで定義する`Message`を生成する責務を持つ。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime, timedelta

# Re:/Fwd: などの返信・転送プレフィックスを除去する正規表現。
# 日本語環境の "RE:", "FW:", "返信:", "転送:" なども対象にする。
_PREFIX_RE = re.compile(
    r"^\s*(re|fwd?|aw|wg|返信|転送|fw)\s*(\[\d+\])?\s*[:：]\s*",
    re.IGNORECASE,
)


def normalize_subject(subject: str) -> str:
    """件名から Re:/Fwd: 等のプレフィックスを繰り返し除去する。

    Args:
        subject: 生の件名。

    Returns:
        正規化済みの件名（前後空白も除去）。
    """
    s = subject or ""
    while True:
        new = _PREFIX_RE.sub("", s)
        if new == s:
            break
        s = new
    return s.strip()


def normalize_email(addr: str | None) -> str:
    """メールアドレスを比較可能な形（小文字・前後空白除去）へ正規化する。

    Args:
        addr: 生のアドレス文字列。`None`可。

    Returns:
        正規化済みアドレス。空なら空文字列。
    """
    return (addr or "").strip().lower()


@dataclass
class Message:
    """1通のメールを表す正規化済みモデル。

    `is_to_me` / `is_cc_me` / `is_from_me` は「自分」の判定結果。生成側で
    `resolve_membership` を呼んで埋めるか、生成時に直接セットする。
    """

    entry_id: str
    store_id: str
    conversation_id: str
    subject: str
    sender_email: str
    sender_name: str
    to_list: list[str] = field(default_factory=list)
    cc_list: list[str] = field(default_factory=list)
    received_time: datetime = field(default_factory=datetime.now)
    body_preview: str = ""
    body_html: str = ""
    unread: bool = False
    importance: int = 1
    folder: str = ""
    is_to_me: bool = False
    is_cc_me: bool = False
    is_from_me: bool = False

    def __post_init__(self) -> None:
        self.sender_email = normalize_email(self.sender_email)
        self.to_list = [normalize_email(a) for a in self.to_list if a]
        self.cc_list = [normalize_email(a) for a in self.cc_list if a]

    @property
    def subject_norm(self) -> str:
        """Re:/Fwd: 除去後の件名。"""
        return normalize_subject(self.subject)

    def resolve_membership(self, my_addresses: set[str]) -> None:
        """`my_addresses` を基準に自分のTo/CC/送信者フラグを確定する。

        To/CCの両方に含まれる場合はToを優先する（自分宛の用件とみなす）。

        Args:
            my_addresses: 自分のSMTPアドレス集合（小文字正規化済み）。
        """
        self.is_from_me = self.sender_email in my_addresses
        self.is_to_me = any(a in my_addresses for a in self.to_list)
        self.is_cc_me = (not self.is_to_me) and any(
            a in my_addresses for a in self.cc_list
        )


# レーン定数。文字列の散らばりを防ぐ。
LANE_RED = "red"  # 🔴 自分ボール
LANE_AMBER = "amber"  # 🟠 活発
LANE_BLUE = "blue"  # 🔵 共有(FYI)
LANE_GRAY = "gray"  # ⚪ 静か

LANE_ORDER = [LANE_RED, LANE_AMBER, LANE_BLUE, LANE_GRAY]
LANE_LABELS = {
    LANE_RED: "🔴 自分ボール",
    LANE_AMBER: "🟠 活発",
    LANE_BLUE: "🔵 共有(FYI)",
    LANE_GRAY: "⚪ 静か",
}


@dataclass
class Conversation:
    """`ConversationID` で束ねた会話。triageが消費する集約モデル。

    triageが参照する主フィールドは `last_from_me` / `i_am_to` /
    `i_am_cc_only` / `velocity_recent`。`lane` / `lane_reason` は仕分け後に
    書き込まれる。
    """

    conversation_id: str
    subject_norm: str
    messages: list[Message] = field(default_factory=list)
    participants: list[str] = field(default_factory=list)
    last_received: datetime | None = None
    last_sender_email: str = ""
    last_from_me: bool = False
    any_unread: bool = False
    i_am_to: bool = False
    i_am_cc_only: bool = False
    velocity_recent: int = 0
    lane: str = LANE_GRAY
    lane_reason: str = ""

    @property
    def last_from_other(self) -> bool:
        """最後の発言が自分以外かどうか（🔴判定のゲート）。"""
        return not self.last_from_me

    @property
    def participant_count(self) -> int:
        """会話の参加者数。"""
        return len(self.participants)
