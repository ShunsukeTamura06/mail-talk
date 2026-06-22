"""返信下書き生成（CLAUDE.md §5）。

常に全員返信（ReplyAll）の下書きを **Outlookで開くだけ**。`.Send()` は絶対に
呼ばない。最終送信は人間が目視確認して押す。返信対象は会話内の最新メール。

実際のCOM操作は `outlook_client.Win32OutlookSource.open_reply_draft` に委譲し、
ここでは「会話 → 最新メールの (entry_id, store_id) 解決」を担う。
"""

from __future__ import annotations

from .db import Database
from .notify import notify_user
from .source import OutlookSource


def open_reply_for_conversation(
    db: Database,
    source: OutlookSource,
    conversation_id: str,
    body_text: str,
) -> bool:
    """会話の最新メールに対する全員返信の下書きをOutlookで開く。

    Args:
        db: メールキャッシュDB。
        source: メール供給元（COM実装またはFake）。
        conversation_id: 返信対象の会話ID。
        body_text: ユーザーが入力した本文（プレーンテキスト）。

    Returns:
        下書きを開けたら True。対象メールが見つからなければ False。
    """
    messages = db.messages_for_conversation(conversation_id)
    if not messages:
        notify_user("warn", "返信対象のメールが見つかりませんでした。")
        return False

    # 会話内の最新メールを対象にする（messages_for_conversationは古い順）。
    latest = messages[-1]
    if not latest.entry_id:
        notify_user("warn", "返信対象のEntryIDが取得できませんでした。")
        return False

    source.open_reply_draft(latest.entry_id, latest.store_id, body_text)
    return True
