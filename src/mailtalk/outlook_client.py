"""win32com による Outlook 取得層（Windows実機専用）。

このモジュールは import 時には win32com を読み込まない（macOSでも import 可能）。
実際のCOM呼び出しは `connect()` 以降で行う。COM依存をこのファイルに閉じ込め、
上位層は `models.Message` と `source.OutlookSource` のみに依存する。

精度の肝は CLAUDE.md §9（自分の特定 / To・CC判別 / 最後の発言 / 例外処理）。
ここでEX(Exchange DN)→SMTP変換を確実に行うことが🔴判定の正確さを決める。
"""

from __future__ import annotations

import time
from datetime import datetime
from typing import Iterator

from .models import Message, normalize_email
from .notify import log_debug, notify_user

# Recipients.Type の定数（win32com非依存で持っておく）。
_OL_TO = 1
_OL_CC = 2
_OL_BCC = 3

# 取得対象フォルダ。olFolderInbox=6, olFolderSentMail=5。
# Sent Itemsも読むのは「最後の発言が自分か（返信済みか）」を正しく判定するため
# （受信トレイだけだと返信済みでも🔴のまま＝痛み④を取りこぼす）。
_OL_FOLDER_INBOX = 6
_OL_FOLDER_SENTMAIL = 5

# MAPIプロパティタグ（PropertyAccessor用フォールバック）。
_PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
_PR_SENDER_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5D01001E"


def _looks_like_smtp(addr: str) -> bool:
    """SMTPアドレスらしい文字列か（EXのDN "/o=..." を除外する）。"""
    return "@" in addr and not addr.startswith("/")


def _coerce_time(mail) -> datetime:
    """メールの時刻を取得する。ReceivedTime優先、無ければSentOn（送信済み対策）。

    Args:
        mail: COMの`MailItem`。

    Returns:
        naiveなdatetime（他レイヤーと揃える）。両方失敗時は現在時刻。
    """
    for attr in ("ReceivedTime", "SentOn"):
        try:
            t = getattr(mail, attr)
            return datetime(t.year, t.month, t.day, t.hour, t.minute, t.second)
        except Exception:  # noqa: BLE001
            continue
    return datetime.now()


def _resolve_smtp(address_entry) -> str:
    """AddressEntry から SMTP アドレスを解決する（CLAUDE.md §9b）。

    EXユーザーは GetExchangeUser → PropertyAccessor(PR_SMTP_ADDRESS) の順で
    試し、それでも解決できなければ Address（DNのことがある）を返す。返り値が
    SMTPらしいかの最終判定は呼び出し側で `_looks_like_smtp` で行う。

    Args:
        address_entry: COMの`AddressEntry`オブジェクト。

    Returns:
        小文字正規化済みアドレス。解決不能時はDN文字列または空。
    """
    if address_entry is None:
        return ""
    try:
        if getattr(address_entry, "Type", "") == "EX":
            try:
                exu = address_entry.GetExchangeUser()
                if exu is not None:
                    smtp = normalize_email(exu.PrimarySmtpAddress)
                    if _looks_like_smtp(smtp):
                        return smtp
            except Exception:  # noqa: BLE001
                pass
            try:
                smtp = normalize_email(
                    address_entry.PropertyAccessor.GetProperty(_PR_SMTP_ADDRESS)
                )
                if _looks_like_smtp(smtp):
                    return smtp
            except Exception:  # noqa: BLE001
                pass
            # 配布リスト等はSMTPを持たないことがある。DN等をそのまま返す。
        return normalize_email(getattr(address_entry, "Address", ""))
    except Exception as exc:  # noqa: BLE001 - COMの多様な例外を握る
        log_debug(f"resolve_smtp失敗: {exc!r}")
        return ""


def _sender_smtp(mail) -> str:
    """メールの送信者SMTPを解決する（EX形式DNを正しくSMTPへ、§9a）。

    GetExchangeUser → PropertyAccessor(PR_SENDER_SMTP_ADDRESS) →
    SenderEmailAddress の順にフォールバックする。

    Args:
        mail: COMの`MailItem`。

    Returns:
        小文字正規化済み送信者SMTPアドレス。
    """
    try:
        if getattr(mail, "SenderEmailType", "") == "EX":
            try:
                smtp = normalize_email(
                    mail.Sender.GetExchangeUser().PrimarySmtpAddress
                )
                if _looks_like_smtp(smtp):
                    return smtp
            except Exception:  # noqa: BLE001
                pass
            try:
                smtp = normalize_email(
                    mail.PropertyAccessor.GetProperty(_PR_SENDER_SMTP_ADDRESS)
                )
                if _looks_like_smtp(smtp):
                    return smtp
            except Exception:  # noqa: BLE001
                pass
            return normalize_email(getattr(mail, "SenderEmailAddress", ""))
        return normalize_email(getattr(mail, "SenderEmailAddress", ""))
    except Exception as exc:  # noqa: BLE001
        log_debug(f"sender_smtp失敗: {exc!r}")
        return ""


class Win32OutlookSource:
    """起動中のOutlookデスクトップに接続するメール供給元。"""

    def __init__(self) -> None:
        self._app = None
        self._ns = None
        self._my_addrs: set[str] | None = None

    # -- 接続 ---------------------------------------------------------------

    def connect(self, retries: int = 30, wait_seconds: float = 2.0) -> None:
        """Outlookへ接続する。未起動ならリトライしながら待つ。

        Args:
            retries: 接続リトライ回数。
            wait_seconds: リトライ間隔（秒）。

        Raises:
            RuntimeError: リトライ上限まで接続できなかった場合。
        """
        import pythoncom  # noqa: F401 - スレッド内COM初期化のため
        import win32com.client

        pythoncom.CoInitialize()
        last_exc: Exception | None = None
        for attempt in range(1, retries + 1):
            try:
                self._app = win32com.client.Dispatch("Outlook.Application")
                self._ns = self._app.GetNamespace("MAPI")
                # アクセスを試行して実際に使えるか確認。
                _ = self._ns.CurrentUser
                notify_user("info", "Outlookに接続しました。")
                return
            except Exception as exc:  # noqa: BLE001
                last_exc = exc
                notify_user(
                    "warn",
                    "Outlookが起動していません。起動すると読み込みを始めます。",
                    detail=f"接続試行 {attempt}/{retries}: {exc!r}",
                )
                time.sleep(wait_seconds)
        raise RuntimeError(f"Outlookへ接続できませんでした: {last_exc!r}")

    def _require_ns(self):
        if self._ns is None:
            self.connect()
        return self._ns

    # -- OutlookSource 契約 -------------------------------------------------

    def my_addresses(self) -> set[str]:
        """自分のSMTPアドレス集合を返す（CLAUDE.md §9a）。"""
        if self._my_addrs is not None:
            return self._my_addrs
        ns = self._require_ns()
        addrs: set[str] = set()
        try:
            exu = ns.CurrentUser.AddressEntry.GetExchangeUser()
            if exu is not None:
                addrs.add(normalize_email(exu.PrimarySmtpAddress))
        except Exception as exc:  # noqa: BLE001
            log_debug(f"CurrentUserのSMTP解決失敗: {exc!r}")
        try:
            for acc in ns.Session.Accounts:
                smtp = getattr(acc, "SmtpAddress", None)
                if smtp:
                    addrs.add(normalize_email(smtp))
        except Exception as exc:  # noqa: BLE001
            log_debug(f"Accounts列挙失敗: {exc!r}")
        if not addrs:
            # 自分のアドレスが1つも取れないと is_from_me/is_to_me が全滅し、
            # 🔴判定の再現率がゼロになる。黙って続けず明示的に失敗させる（§7）。
            notify_user(
                "error",
                "あなたのメールアドレスを特定できませんでした。Outlookのアカウント設定を確認してください。",
                detail="my_addresses() が空。CurrentUser/Accounts いずれもSMTP解決不可。",
            )
            raise RuntimeError("自分のSMTPアドレスを特定できませんでした。")
        self._my_addrs = addrs
        log_debug(f"自分のアドレス: {sorted(addrs)}")
        return addrs

    def iter_messages(
        self, since: datetime | None = None, limit: int | None = None
    ) -> Iterator[Message]:
        """受信トレイ＋送信済みのメールを新しい順で列挙する（CLAUDE.md §9d）。

        送信済みも読むことで「最後の発言が自分か（返信済みか）」を正しく判定する。

        Args:
            since: この時刻より後のメールのみ（差分同期用）。
            limit: 取得上限（全フォルダ合計）。

        Yields:
            正規化・メンバーシップ解決済みの `Message`。
        """
        ns = self._require_ns()
        my = self.my_addresses()

        count = 0
        for folder_id in (_OL_FOLDER_INBOX, _OL_FOLDER_SENTMAIL):
            try:
                folder = ns.GetDefaultFolder(folder_id)
            except Exception as exc:  # noqa: BLE001 - 1フォルダ失敗で全体を止めない
                log_debug(f"フォルダ取得失敗 id={folder_id}: {exc!r}")
                continue
            for msg in self._iter_folder(folder, my, since):
                yield msg
                count += 1
                if limit is not None and count >= limit:
                    return

    def _iter_folder(
        self, folder, my: set[str], since: datetime | None
    ) -> Iterator[Message]:
        """1フォルダ内のメールを新しい順で列挙する（列挙の例外も握る、§9d）。"""
        folder_name = getattr(folder, "Name", "")
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)  # 新しい順
            if since is not None:
                fmt = since.strftime("%m/%d/%Y %H:%M %p")
                try:
                    items = items.Restrict(f"[ReceivedTime] > '{fmt}'")
                except Exception as exc:  # noqa: BLE001
                    log_debug(f"Restrict失敗（全件にフォールバック）: {exc!r}")
        except Exception as exc:  # noqa: BLE001
            log_debug(f"Items取得失敗 folder={folder_name}: {exc!r}")
            return

        try:
            item = items.GetFirst()
        except Exception as exc:  # noqa: BLE001 - GetFirstで死なない
            log_debug(f"GetFirst失敗 folder={folder_name}: {exc!r}")
            return

        while item is not None:
            try:
                if getattr(item, "Class", None) == 43:  # olMail
                    msg = self._to_message(item, my, folder=folder_name)
                    if msg is not None:
                        yield msg
            except Exception as exc:  # noqa: BLE001 - 1通の失敗で全体を止めない
                log_debug(f"メール変換失敗（スキップ）: {exc!r}")
            try:
                item = items.GetNext()  # GetNextの例外でも全体を止めない（§9d）
            except Exception as exc:  # noqa: BLE001
                log_debug(f"GetNext失敗 folder={folder_name}: {exc!r}")
                break

    def _to_message(self, mail, my: set[str], folder: str) -> Message | None:
        """COMの`MailItem`を`Message`へ正規化する。"""
        to_list: list[str] = []
        cc_list: list[str] = []
        to_unresolved = False
        try:
            for r in mail.Recipients:
                smtp = _resolve_smtp(r.AddressEntry)
                rtype = getattr(r, "Type", 0)
                if rtype == _OL_TO:
                    if _looks_like_smtp(smtp):
                        to_list.append(smtp)
                    else:
                        # 解決不能なTo（配布リスト/EX変換失敗）。§9「迷ったら🔴寄り」
                        # のため、自分宛か判別不能としてフラグを立てる。
                        to_unresolved = True
                        log_debug(f"To解決不能: {smtp!r}")
                elif rtype == _OL_CC and _looks_like_smtp(smtp):
                    cc_list.append(smtp)
        except Exception as exc:  # noqa: BLE001
            log_debug(f"Recipients解決失敗: {exc!r}")
            to_unresolved = True  # 解決処理ごと失敗→安全側に倒す

        received_dt = _coerce_time(mail)

        msg = Message(
            entry_id=getattr(mail, "EntryID", ""),
            store_id=getattr(getattr(mail, "Parent", None), "StoreID", "") or "",
            conversation_id=getattr(mail, "ConversationID", "") or "",
            subject=getattr(mail, "Subject", "") or "",
            sender_email=_sender_smtp(mail),
            sender_name=getattr(mail, "SenderName", "") or "",
            to_list=to_list,
            cc_list=cc_list,
            received_time=received_dt,
            body_preview=(getattr(mail, "Body", "") or "")[:500],
            body_html=getattr(mail, "HTMLBody", "") or "",
            unread=bool(getattr(mail, "UnRead", False)),
            importance=int(getattr(mail, "Importance", 1) or 1),
            folder=folder,
            to_unresolved=to_unresolved,
        )
        msg.resolve_membership(my)
        return msg

    def open_reply_draft(self, entry_id: str, store_id: str, body_text: str) -> None:
        """全員返信の下書きをOutlookで開く（CLAUDE.md §5）。送信はしない。

        Args:
            entry_id: 返信対象（会話内最新メール）のEntryID。
            store_id: 同メールのStoreID。
            body_text: ユーザーが入力した本文（プレーンテキスト）。
        """
        ns = self._require_ns()
        original = ns.GetItemFromID(entry_id, store_id)
        reply = original.ReplyAll()  # 常に全員返信
        _ = reply.GetInspector  # 署名+引用元を HTMLBody に挿入させる
        base = reply.HTMLBody
        typed = (body_text or "").replace("\n", "<br>")
        reply.HTMLBody = f"<div>{typed}</div>{base}"
        reply.Display()  # 作成ウィンドウを開くだけ。.Send() は絶対に呼ばない
        notify_user("info", "返信の下書きをOutlookで開きました。内容を確認して送信してください。")
