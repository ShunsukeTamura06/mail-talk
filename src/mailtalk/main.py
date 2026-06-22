"""FastAPIアプリ（CLAUDE.md §10 ステップ3）。

ローカル限定（127.0.0.1）で起動し、チャットUIを配信＋APIを提供する。
起動時にバックグラウンド同期を開始し、UIは準備中でも進捗を見られる。

エンドポイント:
- GET  /                              … チャットUI（static/index.html）
- GET  /api/status                    … 同期状態・進捗・ユーザー通知
- GET  /api/conversations[?lane=red]  … 会話一覧（レーン別）
- GET  /api/conversations/{cid}/messages … 吹き出し用メール（時系列）
- POST /api/conversations/{cid}/reply … 全員返信の下書きをOutlookで開く
- POST /api/sync                      … 同期を手動トリガ（{"full": true}で全件）
- POST /api/diagnostics               … 診断バンドル(zip)を作成しパスを返す
"""

from __future__ import annotations

from contextlib import asynccontextmanager

from fastapi import FastAPI
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from .db import Database
from .notify import get_user_messages
from .paths import static_dir
from .reply import open_reply_for_conversation
from .source import get_default_source
from .sync import SyncManager

# ローカル限定の待受設定。
HOST = "127.0.0.1"
PORT = 8765

_STATIC_DIR = static_dir()

# アプリ全体で共有する状態。
_db = Database()
_source = get_default_source()
_sync = SyncManager(_db, _source)


@asynccontextmanager
async def lifespan(app: FastAPI):
    """起動時にバックグラウンド同期を開始する。"""
    _sync.start_background(full=False)
    yield


app = FastAPI(title="MailTalk", lifespan=lifespan)


class ReplyRequest(BaseModel):
    """返信下書きリクエストのボディ。"""

    body: str = ""


class SyncRequest(BaseModel):
    """同期トリガのボディ。"""

    full: bool = False


@app.get("/api/status")
def api_status() -> dict:
    """同期状態・進捗・直近のユーザー通知を返す。"""
    return {"sync": _sync.status(), "messages": get_user_messages(limit=20)}


@app.get("/api/conversations")
def api_conversations(lane: str | None = None) -> dict:
    """会話一覧を返す（lane指定でレーン別）。"""
    return {"conversations": _db.conversations(lane=lane)}


@app.get("/api/conversations/{conversation_id}/messages")
def api_conversation_messages(conversation_id: str) -> dict:
    """会話のメールを吹き出し表示用に時系列で返す。"""
    msgs = _db.messages_for_conversation(conversation_id)
    return {
        "messages": [
            {
                "entry_id": m.entry_id,
                "sender_name": m.sender_name,
                "sender_email": m.sender_email,
                "is_from_me": m.is_from_me,
                "received_time": m.received_time.isoformat(),
                "subject": m.subject,
                "body_preview": m.body_preview,
                "body_html": m.body_html,
                "unread": m.unread,
            }
            for m in msgs
        ]
    }


@app.post("/api/conversations/{conversation_id}/reply")
def api_reply(conversation_id: str, req: ReplyRequest) -> JSONResponse:
    """全員返信の下書きをOutlookで開く（送信はしない）。"""
    ok = open_reply_for_conversation(_db, _source, conversation_id, req.body)
    if not ok:
        return JSONResponse({"ok": False}, status_code=404)
    return JSONResponse({"ok": True})


@app.post("/api/sync")
def api_sync(req: SyncRequest) -> dict:
    """同期を手動でトリガする。"""
    _sync.start_background(full=req.full)
    return {"ok": True, "status": _sync.status()}


@app.post("/api/diagnostics")
def api_diagnostics(redact: bool = False) -> dict:
    """診断バンドル(zip)を作成し、保存先パスを返す（M端末からの持ち帰り用）。

    Args:
        redact: Trueでアドレス・件名をマスクして収集する。
    """
    from .diagnostics import collect_diagnostics

    path = collect_diagnostics(_source, redact=redact)
    return {"ok": True, "path": str(path)}


@app.get("/")
def index() -> FileResponse:
    """チャットUIを返す。"""
    return FileResponse(str(_STATIC_DIR / "index.html"))


# 静的アセット（将来のCSS/JS分割に備える）。
if _STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(_STATIC_DIR)), name="static")


def main() -> None:
    """uvicornでローカル限定起動し、既定ブラウザでUIを開く。

    M端末ではexeをダブルクリックするだけで使えるよう、起動後に自動で
    ブラウザを開く。待受は127.0.0.1固定（外部公開しない）。

    `MailTalk.exe --diagnostics [--redact]` のように起動した場合はサーバを
    立てず、診断バンドル(zip)だけ作って終了する（M端末でPython無しでも
    持ち帰り用zipを作れるようにするため）。
    """
    import sys

    from .config import ensure_config_file, get_config, reset_config

    # 既定の設定ファイルを用意し、最新を読み込む（M端末で編集可能にする）。
    ensure_config_file()
    reset_config()
    cfg = get_config()

    if "--diagnostics" in sys.argv:
        from .diagnostics import collect_diagnostics

        redact = "--redact" in sys.argv
        path = collect_diagnostics(_source, redact=redact)
        print(f"\n診断バンドルを作成しました。このPCへ持ち帰ってください:\n  {path}")
        return

    import threading
    import webbrowser

    import uvicorn

    url = f"http://{cfg.host}:{cfg.port}"
    # サーバ起動直後にブラウザを開く（設定で無効化可）。
    if cfg.open_browser:
        threading.Timer(1.5, lambda: webbrowser.open(url)).start()
    uvicorn.run(app, host=cfg.host, port=cfg.port)


if __name__ == "__main__":
    main()
