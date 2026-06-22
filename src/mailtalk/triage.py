"""レーン仕分け（最重要ロジック）。

会話を点数で並べるのではなく、上から順に条件を当て、最初に当たったレーンで
確定する「優先順位カスケード」。🔴を必ず最初に判定する。

設計判断（CLAUDE.md §3 より、絶対に守る）:
- 🔴を活発より先に評価する（活発さは🔴を押しのけない）。
- 🔴のゲートは既読/未読ではなく「最後が相手か」という返信状態。
- エラーコストは非対称。迷ったら🔴に倒す（再現率優先）。
"""

from __future__ import annotations

from .models import (
    LANE_AMBER,
    LANE_BLUE,
    LANE_GRAY,
    LANE_RED,
    Conversation,
)

# 直近24-48hの件数しきい値。実データを見て調整可能にする。
ACTIVE_THRESHOLD = 6


def classify(conv: Conversation) -> tuple[str, str]:
    """会話を4レーンのいずれかに仕分けし、理由文とともに返す。

    優先順位カスケード（最初に当たったレーンで確定）:
        1. 🔴 自分ボール: 最後が相手の発言 かつ 自分がToに入っている
        2. 🟠 活発: （🔴でない）かつ 直近の往復が多い
        3. 🔵 共有(FYI): （上2つでない）かつ 自分がCCのみ
        4. ⚪ 静か: 残り全部

    Args:
        conv: 集約済みの会話。`last_from_other` / `i_am_to` /
            `velocity_recent` / `i_am_cc_only` が埋まっていること。

    Returns:
        `(lane, lane_reason)` のタプル。`lane` は `models` のレーン定数、
        `lane_reason` はUI表示用の人間可読な理由文。
    """
    # 1. 🔴 自分ボール（最初に判定）
    if conv.last_from_other and conv.i_am_to:
        return LANE_RED, "Toに自分・最後が相手の発言（あなたが返す番）"

    # （将来）CC埋もれ依頼の救済:
    # if conv.last_from_other and conv.has_request_to_me:
    #     return LANE_RED, "CCだが本文であなたに依頼あり"

    # 2. 🟠 活発
    if conv.velocity_recent >= ACTIVE_THRESHOLD:
        return LANE_AMBER, f"直近{conv.velocity_recent}件・活発にやり取り中"

    # 3. 🔵 共有(FYI)
    if conv.i_am_cc_only:
        return LANE_BLUE, "CCのみ・あなた宛の用件なし"

    # 4. ⚪ 静か
    return LANE_GRAY, "動きなし"


def classify_into(conv: Conversation) -> Conversation:
    """`classify` の結果を会話オブジェクトへ書き込んで返す。

    Args:
        conv: 仕分け対象の会話（破壊的に更新される）。

    Returns:
        `lane` / `lane_reason` を更新した同一オブジェクト。
    """
    conv.lane, conv.lane_reason = classify(conv)
    return conv
