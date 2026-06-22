"""コンソール検証ツール（CLAUDE.md §10 ステップ1の成果物）。

「会話一覧を読み込んで、各会話がどのレーンか／最後が誰の発言か／自分はTo・CC
どっちか、をコンソールに吐く」。実データ（Windows実機）で仕分け精度を確認する
ための最小ツール。macOSではFakeデータで動作する。

使い方:
    uv run mailtalk                 # 会話一覧をレーン別に表示
    uv run mailtalk --limit 200     # 取得上限を指定
    uv run mailtalk --diagnostics   # 診断バンドル（持ち帰り最小化）を logs/ へ
"""

from __future__ import annotations

import argparse
import platform
import sys
import time
from datetime import datetime

from .aggregate import build_conversations
from .models import LANE_LABELS, LANE_ORDER
from .notify import log_debug, notify_user
from .source import get_default_source
from .triage import classify_into


def _load_classified(limit: int | None) -> list:
    """供給元からメールを取得し、会話に集約してレーン仕分けする。

    Args:
        limit: 取得上限。

    Returns:
        仕分け済み会話のリスト（最終受信が新しい順）。
    """
    source = get_default_source()
    # 実機(win32)では接続待ちが入る。Fakeはno-op。
    if hasattr(source, "connect"):
        source.connect()

    t0 = time.perf_counter()
    messages = list(source.iter_messages(limit=limit))
    log_debug(f"取得 {len(messages)}件 / {time.perf_counter() - t0:.2f}s")

    convs = build_conversations(messages)
    for c in convs:
        classify_into(c)
    return convs


def _print_lanes(convs: list) -> None:
    """会話一覧をレーン別にコンソール出力する。"""
    by_lane: dict[str, list] = {lane: [] for lane in LANE_ORDER}
    for c in convs:
        by_lane[c.lane].append(c)

    print()
    print(f"=== MailTalk 仕分け結果（全{len(convs)}会話）===")
    for lane in LANE_ORDER:
        items = by_lane[lane]
        print(f"\n{LANE_LABELS[lane]}  ({len(items)}件)")
        print("-" * 60)
        for c in items:
            last = "自分" if c.last_from_me else c.last_sender_email
            role = "To:自分" if c.i_am_to else ("CCのみ" if c.i_am_cc_only else "-")
            unread = "未読" if c.any_unread else "  "
            print(
                f"  [{unread}] {c.subject_norm[:32]:<32} "
                f"| 最後={last:<22} | {role:<7} "
                f"| 人数{c.participant_count} 直近{c.velocity_recent}"
            )
            print(f"        └ 理由: {c.lane_reason}")
    print()


def _run_diagnostics(limit: int | None) -> None:
    """診断情報を logs/ に出力する（会社端末からの持ち帰り最小化用）。

    シークレット（アドレスの実値以外の機密）は出さない。
    """
    notify_user("info", "診断モードを開始します。")
    log_debug("=== 診断バンドル ===")
    log_debug(f"platform={platform.platform()} python={sys.version.split()[0]}")
    log_debug(f"started_at={datetime.now().isoformat()}")
    t0 = time.perf_counter()
    convs = _load_classified(limit)
    counts = {lane: 0 for lane in LANE_ORDER}
    for c in convs:
        counts[c.lane] += 1
    log_debug(f"会話数={len(convs)} 内訳={counts} 所要={time.perf_counter() - t0:.2f}s")
    notify_user(
        "info",
        f"診断完了。{len(convs)}会話を仕分けしました（内訳 {counts}）。logs/app.log を確認してください。",
    )


def main() -> None:
    """CLIエントリポイント。"""
    parser = argparse.ArgumentParser(description="MailTalk コンソール検証ツール")
    parser.add_argument("--limit", type=int, default=None, help="取得上限件数")
    parser.add_argument(
        "--diagnostics", action="store_true", help="診断バンドルを logs/ へ出力"
    )
    args = parser.parse_args()

    if args.diagnostics:
        _run_diagnostics(args.limit)
        return

    convs = _load_classified(args.limit)
    _print_lanes(convs)


if __name__ == "__main__":
    main()
