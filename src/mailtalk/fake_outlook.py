"""開発機（macOS等）用のFakeメール供給元。

`OutlookSource` と同じインターフェースを提供し、4レーンを網羅する合成データを
返す。これによりCOM無しで「取得→集約→仕分け→DB→表示」のパイプライン全体を
検証できる（CLAUDE.md §13 環境の二重性）。
"""

from __future__ import annotations

from datetime import datetime, timedelta
from typing import Iterator

from .models import Message
from .notify import notify_user

ME = "me@bank.co.jp"
ME_NAME = "田村 俊輔"


def _msg(
    eid: str,
    conv: str,
    subject: str,
    sender_email: str,
    sender_name: str,
    to: list[str],
    cc: list[str],
    minutes_ago: int,
    *,
    unread: bool = False,
    body: str = "",
) -> Message:
    """テスト用に1通のMessageを組み立てる（基準時刻=now）。"""
    m = Message(
        entry_id=eid,
        store_id="FAKESTORE",
        conversation_id=conv,
        subject=subject,
        sender_email=sender_email,
        sender_name=sender_name,
        to_list=to,
        cc_list=cc,
        received_time=datetime.now() - timedelta(minutes=minutes_ago),
        body_preview=body or f"{subject} の本文プレビュー。",
        body_html=f"<p>{body or subject}</p>",
        unread=unread,
    )
    m.resolve_membership({ME})
    return m


def _build_dataset() -> list[Message]:
    """4レーンを網羅する合成メール群を返す。"""
    boss = ("tanaka@bank.co.jp", "田中 部長")
    sato = ("sato@partner.co.jp", "佐藤 様")
    suzuki = ("suzuki@bank.co.jp", "鈴木")
    yamada = ("yamada@bank.co.jp", "山田")
    sysadmin = ("noreply@system.bank.co.jp", "システム通知")

    msgs: list[Message] = []

    # 🔴 自分ボール: 最後が相手・自分がTo。返す番。
    msgs += [
        _msg("r1", "CONV_RED", "【要確認】月次バッチの結果について",
             ME, ME_NAME, [boss[0]], [], 200),
        _msg("r2", "CONV_RED", "RE: 【要確認】月次バッチの結果について",
             boss[0], boss[1], [ME], [], 90, unread=True,
             body="田村さん、添付の件、確認して折り返しお願いします。"),
    ]

    # 🟠 活発: 🔴でない（最後が自分）かつ直近の往復が多い。
    base = "CONV_AMBER"
    senders = [sato, ("me", "self"), sato, suzuki, sato, ("me", "self")]
    for i in range(6):
        who = senders[i]
        from_me = who[0] == "me"
        msgs.append(
            _msg(
                f"a{i}", base, "障害対応の進捗共有",
                ME if from_me else who[0],
                ME_NAME if from_me else who[1],
                [sato[0]] if from_me else [ME, suzuki[0]],
                [] if from_me else [yamada[0]],
                40 - i * 5,
                body="進捗を共有します。",
            )
        )

    # 🔵 共有(FYI): 自分はCCのみ、最後が相手。
    msgs += [
        _msg("b1", "CONV_BLUE", "[共有] 新リリースのお知らせ",
             yamada[0], yamada[1], [suzuki[0]], [ME], 300,
             body="関係各位、リリース予定を共有します。"),
        _msg("b2", "CONV_BLUE", "RE: [共有] 新リリースのお知らせ",
             suzuki[0], suzuki[1], [yamada[0]], [ME], 280),
    ]

    # ⚪ 静か: 動きなし。最後が自分・古い・自分宛でもない単発。
    msgs += [
        _msg("g1", "CONV_GRAY", "先週の議事録",
             ME, ME_NAME, [suzuki[0]], [], 5000,
             body="議事録を送ります。"),
        _msg("g2", "CONV_GRAY2", "システムメンテナンスのお知らせ",
             sysadmin[0], sysadmin[1], ["all@bank.co.jp"], [], 6000,
             body="定期メンテナンスを実施します。"),
    ]
    return msgs


class FakeOutlookSource:
    """合成データを返す開発用のメール供給元。"""

    def __init__(self) -> None:
        self._messages = _build_dataset()

    def connect(self, retries: int = 0, wait_seconds: float = 0.0) -> None:
        """互換用のno-op接続。"""
        notify_user("info", "（開発モード）Fakeデータに接続しました。")

    def my_addresses(self) -> set[str]:
        """自分のアドレス集合を返す。"""
        return {ME}

    def iter_messages(
        self,
        since: datetime | None = None,
        before: datetime | None = None,
        limit: int | None = None,
    ) -> Iterator[Message]:
        """合成メールを新しい順で列挙する。

        Args:
            since: この時刻より後のみ（差分同期の挙動を模倣）。
            before: この時刻より前のみ（バックフィルの挙動を模倣）。
            limit: 取得上限。
        """
        ordered = sorted(self._messages, key=lambda m: m.received_time, reverse=True)
        count = 0
        for m in ordered:
            if since is not None and m.received_time <= since:
                continue
            if before is not None and m.received_time >= before:
                continue
            yield m
            count += 1
            if limit is not None and count >= limit:
                return

    def open_reply_draft(self, entry_id: str, store_id: str, body_text: str) -> None:
        """下書きを開く操作を模倣する（実際には何もしない）。"""
        notify_user(
            "info",
            "（開発モード）返信下書きを開いたつもりです。実機ではOutlookが開きます。",
            detail=f"entry_id={entry_id} body_len={len(body_text or '')}",
        )
