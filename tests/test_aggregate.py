"""aggregate（集約）とモデル正規化のテスト。"""

from __future__ import annotations

from datetime import datetime, timedelta

from mailtalk.aggregate import build_conversations
from mailtalk.models import Message, normalize_subject

ME = "me@bank.co.jp"
OTHER = "other@partner.co.jp"


def _m(eid, conv, sender, to, cc, minutes_ago, now) -> Message:
    m = Message(
        entry_id=eid,
        store_id="S",
        conversation_id=conv,
        subject="件名",
        sender_email=sender,
        sender_name=sender,
        to_list=to,
        cc_list=cc,
        received_time=now - timedelta(minutes=minutes_ago),
    )
    m.resolve_membership({ME})
    return m


def test_normalize_subject_strips_prefixes():
    assert normalize_subject("RE: Fwd: 重要な件") == "重要な件"
    assert normalize_subject("返信: お知らせ") == "お知らせ"
    assert normalize_subject("そのまま") == "そのまま"


def test_membership_to_takes_priority_over_cc():
    m = _m("1", "c", OTHER, [ME], [ME], 1, datetime.now())
    assert m.is_to_me is True
    assert m.is_cc_me is False


def test_build_groups_by_conversation_id():
    now = datetime.now()
    msgs = [
        _m("1", "A", ME, [OTHER], [], 30, now),
        _m("2", "A", OTHER, [ME], [], 10, now),
        _m("3", "B", OTHER, [ME], [], 5, now),
    ]
    convs = build_conversations(msgs, now=now)
    assert len(convs) == 2
    # 最終受信が新しい順（Bが先頭）。
    assert convs[0].conversation_id == "B"


def test_last_from_and_signals():
    now = datetime.now()
    msgs = [
        _m("1", "A", ME, [OTHER], [], 30, now),
        _m("2", "A", OTHER, [ME], [], 10, now),
    ]
    conv = build_conversations(msgs, now=now)[0]
    assert conv.last_from_me is False
    assert conv.last_from_other is True
    assert conv.i_am_to is True
    assert conv.velocity_recent == 2


def test_velocity_only_counts_recent_window():
    now = datetime.now()
    msgs = [
        _m("old", "A", OTHER, [ME], [], 60 * 24 * 5, now),  # 5日前
        _m("new", "A", OTHER, [ME], [], 10, now),
    ]
    conv = build_conversations(msgs, now=now)[0]
    assert conv.velocity_recent == 1


def test_cc_only_signal():
    now = datetime.now()
    msgs = [_m("1", "A", OTHER, [OTHER], [ME], 10, now)]
    conv = build_conversations(msgs, now=now)[0]
    assert conv.i_am_to is False
    assert conv.i_am_cc_only is True


def test_signals_use_latest_inbound_not_history():
    # 過去はTo、最新の相手メールではCCのみ → i_am_toは最新基準でFalse。
    now = datetime.now()
    msgs = [
        _m("old", "A", OTHER, [ME], [], 120, now),          # 昔はTo
        _m("new", "A", OTHER, [OTHER], [ME], 10, now),       # 今はCCのみ
    ]
    conv = build_conversations(msgs, now=now)[0]
    assert conv.i_am_to is False
    assert conv.i_am_cc_only is True


def test_latest_inbound_ignores_my_own_last_send():
    # 最新が自分発でも、宛先信号は直近の「相手の発言」で見る。
    now = datetime.now()
    msgs = [
        _m("in", "A", OTHER, [ME], [], 30, now),   # 相手→自分(To)
        _m("out", "A", ME, [OTHER], [], 10, now),  # 自分が返信
    ]
    conv = build_conversations(msgs, now=now)[0]
    assert conv.i_am_to is True          # 直近の相手発言は自分宛
    assert conv.last_from_me is True     # ただし最後は自分→🔴にはならない


def test_unresolved_to_leans_red():
    # 配布リスト宛等でToが解決不能な相手メール → 安全側で i_am_to=True(🔴寄り)。
    now = datetime.now()
    m = Message(
        entry_id="x", store_id="S", conversation_id="A",
        subject="件名", sender_email=OTHER, sender_name=OTHER,
        to_list=[], cc_list=[],
        received_time=now - timedelta(minutes=5),
        to_unresolved=True,
    )
    m.resolve_membership({ME})
    assert m.is_to_me is False
    conv = build_conversations([m], now=now)[0]
    assert conv.i_am_to is True
