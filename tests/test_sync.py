"""同期（コールドスタートのウィンドウ化＋バックフィル）のテスト。"""

from __future__ import annotations

from datetime import datetime, timedelta

from mailtalk import config
from mailtalk.db import Database
from mailtalk.fake_outlook import FakeOutlookSource
from mailtalk.models import LANE_RED
from mailtalk.sync import STATE_READY, SyncManager


def test_fake_before_and_since_filters():
    src = FakeOutlookSource()
    now = datetime.now()
    cutoff = now - timedelta(days=1)
    recent = list(src.iter_messages(since=cutoff))
    old = list(src.iter_messages(before=cutoff))
    all_msgs = list(src.iter_messages())
    # 直近と古いで重複なく全件を分割できる。
    assert len(recent) + len(old) == len(all_msgs)
    assert all(m.received_time > cutoff for m in recent)
    assert all(m.received_time < cutoff for m in old)


def test_cold_start_window_then_backfill():
    # ウィンドウを1日に縮め、Fakeの古いメール(数日前)をバックフィル対象にする。
    config.set_config(config.Config(cold_window_days=1, backfill_old=True))
    db = Database(":memory:")
    mgr = SyncManager(db, FakeOutlookSource())

    mgr.run(full=False)  # コールドパス（同期実行）

    status = mgr.status()
    assert status["state"] == STATE_READY
    # ウィンドウ＋バックフィルで全件入る。
    assert len(db.all_messages()) == len(list(FakeOutlookSource().iter_messages()))
    # 起点とバックフィル完了フラグが立つ。
    assert db.get_state("last_sync_time") is not None
    assert db.get_state("backfill_done") == "1"
    # 仕分け済みで🔴が1件以上。
    assert len(db.conversations(lane=LANE_RED)) >= 1
    db.close()


def test_backfill_disabled_loads_recent_only():
    # backfill_old=False なら直近ウィンドウのみ読み込み、古い分は取り込まない。
    config.set_config(config.Config(cold_window_days=1, backfill_old=False))
    db = Database(":memory:")
    mgr = SyncManager(db, FakeOutlookSource())
    mgr.run(full=False)

    from datetime import datetime, timedelta

    recent = list(FakeOutlookSource().iter_messages(since=datetime.now() - timedelta(days=1)))
    assert len(db.all_messages()) == len(recent)
    assert db.get_state("backfill_done") is None
    assert mgr.status()["state"] == STATE_READY
    db.close()


def test_incremental_after_cold():
    # 2回目以降は last_sync_time を起点に差分同期する（コールド分岐に入らない）。
    config.set_config(config.Config(cold_window_days=1, backfill_old=True))
    db = Database(":memory:")
    mgr = SyncManager(db, FakeOutlookSource())
    mgr.run(full=False)
    first_last = db.get_state("last_sync_time")

    # 差分なし（新着メールが無い）なら件数は変わらず、準備完了に戻る。
    mgr.run(full=False)
    assert mgr.status()["state"] == STATE_READY
    assert db.get_state("last_sync_time") == first_last
    db.close()
