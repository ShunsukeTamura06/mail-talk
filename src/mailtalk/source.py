"""メール取得層の抽象インターフェース。

上位層（aggregate / triage / db / main）はこの`OutlookSource`にのみ依存し、
COM実装(`outlook_client.Win32OutlookSource`)とFake実装
(`fake_outlook.FakeOutlookSource`)を差し替え可能にする。これにより
macOS開発機ではFakeでパイプライン全体を流し、Windows実機ではCOMで動かす。
"""

from __future__ import annotations

import sys
from datetime import datetime
from typing import Iterator, Protocol, runtime_checkable

from .models import Message


@runtime_checkable
class OutlookSource(Protocol):
    """Outlook（または代替）からメールを取得する供給元の契約。"""

    def my_addresses(self) -> set[str]:
        """自分のSMTPアドレス集合（小文字正規化済み）を返す。"""
        ...

    def iter_messages(
        self,
        since: datetime | None = None,
        before: datetime | None = None,
        limit: int | None = None,
    ) -> Iterator[Message]:
        """メールを新しい順で列挙する。

        Args:
            since: この時刻より後のメールのみ（差分同期用）。Noneで下限なし。
            before: この時刻より前のメールのみ（古い分のバックフィル用）。Noneで上限なし。
            limit: 取得上限。Noneで無制限。
        """
        ...

    def open_reply_draft(self, entry_id: str, store_id: str, body_text: str) -> None:
        """全員返信の下書きをOutlookで開く（送信はしない）。"""
        ...


def get_default_source() -> OutlookSource:
    """実行環境に応じた既定のメール供給元を返す。

    Windowsでは`Win32OutlookSource`、それ以外（開発機）では
    `FakeOutlookSource`を返す。これにより同じ上位コードがどちらでも動く。

    Returns:
        `OutlookSource` を満たすインスタンス。
    """
    if sys.platform == "win32":
        from .outlook_client import Win32OutlookSource

        return Win32OutlookSource()

    from .fake_outlook import FakeOutlookSource

    return FakeOutlookSource()
