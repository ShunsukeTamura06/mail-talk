"""変換・集約層: Message群を Conversation へ束ね、triage用の信号を計算する。

COM非依存の純粋ロジック。macOS開発機でFakeデータにより全面検証できる。
"""

from __future__ import annotations

from datetime import datetime, timedelta

from .config import get_config
from .models import Conversation, Message

# 「活発さ」を測る直近ウィンドウ。CLAUDE.md §3 の「直近24-48h」を48hで採用。
RECENT_WINDOW = timedelta(hours=48)


def _participants(messages: list[Message]) -> list[str]:
    """会話の参加者（送信者＋To＋CC）を重複なく出現順で返す。"""
    seen: dict[str, None] = {}
    for m in sorted(messages, key=lambda x: x.received_time):
        for addr in [m.sender_email, *m.to_list, *m.cc_list]:
            if addr and addr not in seen:
                seen[addr] = None
    return list(seen.keys())


def build_conversation(
    conversation_id: str,
    messages: list[Message],
    *,
    now: datetime | None = None,
) -> Conversation:
    """同一会話のMessage群から集約済みの Conversation を構築する。

    Args:
        conversation_id: 会話ID（Outlookの`ConversationID`）。
        messages: この会話に属するメール（メンバーシップ解決済みであること）。
        now: 「活発さ」算定の基準時刻。省略時は現在時刻。テスト用に注入可能。

    Returns:
        triage信号（`last_from_me` / `i_am_to` / `i_am_cc_only` /
        `velocity_recent` 等）を埋めた会話。`lane` は未設定（triageで確定）。
    """
    now = now or datetime.now()
    ordered = sorted(messages, key=lambda m: m.received_time)
    last = ordered[-1]

    # レーン判定の宛先信号は「いま返す相手の発言＝最新の相手メール」基準で計算する。
    # 過去にToだっただけで🔴になるのを防ぐ（最新でCCのみなら🔴にしない）。相手の
    # 発言が無ければ（全部自分発）最後のメールを基準にする（どのみち🔴にならない）。
    inbound = [m for m in ordered if not m.is_from_me]
    ref = inbound[-1] if inbound else last
    # 🔴は「自分が実際にToに居る」を基準にする。解決不能To(配布リスト等)を🔴寄りに
    # 倒すのは既定OFF（解決失敗が多い環境で🔴が暴発するため）。config で有効化可。
    i_am_to = ref.is_to_me or (
        ref.to_unresolved and get_config().red_on_unresolved_to
    )
    i_am_cc = ref.is_cc_me

    # velocity_recentは直近ウィンドウ内の件数。NOTE(#5): 「往復」(送信者交代)では
    # なく単純件数のため、連投も🟠になりうる。🟠は🔴の後段で最低リスクのためMVP許容。
    velocity_recent = sum(1 for m in ordered if m.received_time >= now - RECENT_WINDOW)

    subject_norm = next(
        (m.subject_norm for m in ordered if m.subject_norm),
        ordered[0].subject_norm,
    )

    return Conversation(
        conversation_id=conversation_id,
        subject_norm=subject_norm,
        messages=ordered,
        participants=_participants(ordered),
        last_received=last.received_time,
        last_sender_email=last.sender_email,
        last_from_me=last.is_from_me,
        any_unread=any(m.unread for m in ordered),
        i_am_to=i_am_to,
        i_am_cc_only=(not i_am_to) and i_am_cc,
        velocity_recent=velocity_recent,
    )


def build_conversations(
    messages: list[Message],
    *,
    now: datetime | None = None,
) -> list[Conversation]:
    """Message群を会話単位にグルーピングし、最終受信が新しい順に返す。

    Args:
        messages: 全メール（メンバーシップ解決済み）。
        now: 「活発さ」算定の基準時刻。省略時は現在時刻。

    Returns:
        最終受信時刻の降順に並べた Conversation のリスト。
    """
    now = now or datetime.now()
    buckets: dict[str, list[Message]] = {}
    for m in messages:
        # ConversationIDが空のメールはentry_id単独の会話として扱う。
        key = m.conversation_id or f"__single__:{m.entry_id}"
        buckets.setdefault(key, []).append(m)

    convs = [build_conversation(cid, msgs, now=now) for cid, msgs in buckets.items()]
    convs.sort(key=lambda c: c.last_received or datetime.min, reverse=True)
    return convs
