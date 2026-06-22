"""DB往復とFake供給元→パイプライン結合のテスト（macOSで完結検証）。"""

from __future__ import annotations

from mailtalk.aggregate import build_conversations
from mailtalk.db import Database
from mailtalk.fake_outlook import FakeOutlookSource
from mailtalk.models import LANE_AMBER, LANE_BLUE, LANE_GRAY, LANE_RED
from mailtalk.triage import classify_into


def test_fake_pipeline_covers_all_lanes():
    source = FakeOutlookSource()
    messages = list(source.iter_messages())
    convs = build_conversations(messages)
    for c in convs:
        classify_into(c)
    lanes = {c.lane for c in convs}
    assert {LANE_RED, LANE_AMBER, LANE_BLUE, LANE_GRAY} <= lanes


def test_db_roundtrip_in_memory():
    db = Database(":memory:")
    source = FakeOutlookSource()
    messages = list(source.iter_messages())
    for m in messages:
        db.upsert_email(m)

    convs = build_conversations(db.all_messages())
    for c in convs:
        classify_into(c)
        db.upsert_conversation(c)

    red = db.conversations(lane=LANE_RED)
    assert len(red) >= 1
    assert red[0]["lane"] == LANE_RED
    assert red[0]["lane_reason"]

    # 会話のメール取り出し（吹き出し用）。
    cid = red[0]["conversation_id"]
    msgs = db.messages_for_conversation(cid)
    assert len(msgs) >= 1
    # 古い順で返る。
    assert msgs == sorted(msgs, key=lambda m: m.received_time)
    db.close()


def test_sync_state_kv():
    db = Database(":memory:")
    assert db.get_state("last_sync_time") is None
    db.set_state("last_sync_time", "2026-06-22T10:00:00")
    assert db.get_state("last_sync_time") == "2026-06-22T10:00:00"
    db.close()
