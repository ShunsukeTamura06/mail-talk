"""triage（レーン仕分け）のテスト。CLAUDE.md §3 の設計判断を固定する。"""

from __future__ import annotations

from mailtalk.models import (
    LANE_AMBER,
    LANE_BLUE,
    LANE_GRAY,
    LANE_RED,
    Conversation,
)
from mailtalk.triage import ACTIVE_THRESHOLD, classify


def _conv(**kw) -> Conversation:
    base = dict(conversation_id="c", subject_norm="s")
    base.update(kw)
    return Conversation(**base)


def test_red_when_last_from_other_and_i_am_to():
    c = _conv(last_from_me=False, i_am_to=True)
    lane, reason = classify(c)
    assert lane == LANE_RED
    assert reason


def test_red_beats_active():
    # 激活発でも、最後が相手＆自分がToなら問答無用で🔴。
    c = _conv(last_from_me=False, i_am_to=True, velocity_recent=999)
    assert classify(c)[0] == LANE_RED


def test_red_gate_is_reply_state_not_unread():
    # 既読でも（any_unread=False）最後が相手なら🔴に残る。
    c = _conv(last_from_me=False, i_am_to=True, any_unread=False)
    assert classify(c)[0] == LANE_RED


def test_not_red_when_last_from_me():
    # 最後が自分なら🔴ではない（ボールは相手にある）。
    c = _conv(last_from_me=True, i_am_to=True, velocity_recent=ACTIVE_THRESHOLD)
    assert classify(c)[0] != LANE_RED


def test_amber_when_active_and_not_red():
    c = _conv(last_from_me=True, i_am_to=True, velocity_recent=ACTIVE_THRESHOLD)
    assert classify(c)[0] == LANE_AMBER


def test_blue_when_cc_only_and_not_active():
    c = _conv(
        last_from_me=False,
        i_am_to=False,
        i_am_cc_only=True,
        velocity_recent=ACTIVE_THRESHOLD - 1,
    )
    assert classify(c)[0] == LANE_BLUE


def test_gray_when_nothing_matches():
    c = _conv(last_from_me=True, i_am_to=False, i_am_cc_only=False, velocity_recent=0)
    assert classify(c)[0] == LANE_GRAY


def test_below_threshold_is_not_amber():
    c = _conv(last_from_me=True, i_am_to=True, velocity_recent=ACTIVE_THRESHOLD - 1)
    assert classify(c)[0] != LANE_AMBER
