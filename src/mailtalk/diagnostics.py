"""診断バンドル収集（CLAUDE.md §13 / 全社方針「1回の持ち帰りで原因究明」）。

M端末で `mailtalk --diagnostics` を1コマンド実行すると、🔴判定の検証に必要な
情報を1つのzipにまとめる。会社端末からこのPCへ持ち帰る回数を最小化するのが目的。

含めるもの: 環境情報 / 自分アドレス特定の結果 / 所要時間 / レーン内訳 / 各会話・
各メールのメタデータ（送信者・To/CC判定・to_unresolved・最後が誰か等）/ app.log。
含めないもの: メール本文（body_preview/body_html）。`redact=True` でアドレス・
件名もマスクできる（トークン/APIキー等のシークレットは元々ログに出さない設計）。
"""

from __future__ import annotations

import json
import platform
import sys
import time
import zipfile
from datetime import datetime
from pathlib import Path

from . import __version__
from .aggregate import build_conversations
from .models import LANE_ORDER
from .paths import is_frozen, logs_dir
from .source import OutlookSource, get_default_source
from .triage import classify_into


def _mask_email(addr: str) -> str:
    """アドレスのローカル部をマスクする（ドメインは残す）。"""
    if not addr or "@" not in addr:
        return addr or ""
    local, _, domain = addr.partition("@")
    return f"{local[:2]}***@{domain}"


def collect_diagnostics(
    source: OutlookSource | None = None,
    *,
    limit: int | None = None,
    redact: bool = False,
    out_dir: Path | None = None,
) -> Path:
    """診断情報を収集し、1つのzipにまとめて返す。

    取得・自己特定・列挙の各段で失敗しても中断せず、失敗内容をレポートに記録する
    （失敗自体が最重要の診断材料のため）。

    Args:
        source: メール供給元。省略時は環境に応じた既定。
        limit: 取得上限。
        redact: Trueでアドレス・件名・会話IDをマスクする。
        out_dir: zip出力先。省略時は logs/。

    Returns:
        生成したzipファイルのパス。
    """
    source = source or get_default_source()
    out_dir = Path(out_dir) if out_dir else logs_dir()
    out_dir.mkdir(parents=True, exist_ok=True)

    def a(addr: str) -> str:
        return _mask_email(addr) if redact else addr

    def s(text: str) -> str:
        return f"(件名 len={len(text or '')})" if redact else text

    started = datetime.now()
    errors: list[str] = []

    t0 = time.perf_counter()
    try:
        if hasattr(source, "connect"):
            source.connect()
    except Exception as exc:  # noqa: BLE001 - 接続失敗も診断材料
        errors.append(f"connect: {exc!r}")

    my: set[str] = set()
    try:
        my = source.my_addresses()
    except Exception as exc:  # noqa: BLE001
        errors.append(f"my_addresses: {exc!r}")

    messages = []
    try:
        messages = list(source.iter_messages(limit=limit))
    except Exception as exc:  # noqa: BLE001
        errors.append(f"iter_messages: {exc!r}")
    fetch_sec = time.perf_counter() - t0

    t1 = time.perf_counter()
    convs = build_conversations(messages)
    for c in convs:
        classify_into(c)
    classify_sec = time.perf_counter() - t1

    lane_counts = {lane: 0 for lane in LANE_ORDER}
    for c in convs:
        lane_counts[c.lane] += 1

    conv_dump = []
    for c in convs:
        cid = c.conversation_id
        conv_dump.append(
            {
                "conversation_id": f"conv_{abs(hash(cid)) % 1000000}" if redact else cid,
                "subject_norm": s(c.subject_norm),
                "lane": c.lane,
                "lane_reason": c.lane_reason,
                "last_from_me": c.last_from_me,
                "last_sender": a(c.last_sender_email),
                "i_am_to": c.i_am_to,
                "i_am_cc_only": c.i_am_cc_only,
                "velocity_recent": c.velocity_recent,
                "participant_count": c.participant_count,
                "any_unread": c.any_unread,
                "last_received": c.last_received.isoformat() if c.last_received else None,
                "messages": [
                    {
                        "sender": a(m.sender_email),
                        "is_from_me": m.is_from_me,
                        "is_to_me": m.is_to_me,
                        "is_cc_me": m.is_cc_me,
                        "to_unresolved": m.to_unresolved,
                        "received_time": m.received_time.isoformat(),
                        "folder": m.folder,
                        "unread": m.unread,
                    }
                    for m in c.messages
                ],
            }
        )

    unresolved = sum(1 for c in convs for m in c.messages if m.to_unresolved)

    report = {
        "schema": 1,
        "app_version": __version__,
        "generated_at": started.isoformat(),
        "redacted": redact,
        "errors": errors,
        "environment": {
            "platform": platform.platform(),
            "python": sys.version.split()[0],
            "frozen": is_frozen(),
            "source": type(source).__name__,
        },
        "timing_sec": {
            "fetch": round(fetch_sec, 2),
            "classify": round(classify_sec, 2),
        },
        "self_identification": {
            "my_addresses": [a(x) for x in sorted(my)],
            "count": len(my),
        },
        "summary": {
            "messages": len(messages),
            "conversations": len(convs),
            "lane_counts": lane_counts,
            "to_unresolved_messages": unresolved,
        },
        "conversations": conv_dump,
    }

    stamp = started.strftime("%Y%m%d_%H%M%S")
    zip_path = out_dir / f"mailtalk_diag_{stamp}.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "diagnostics.json", json.dumps(report, ensure_ascii=False, indent=2)
        )
        log_file = logs_dir() / "app.log"
        if log_file.exists():
            z.write(log_file, "app.log")

    return zip_path
