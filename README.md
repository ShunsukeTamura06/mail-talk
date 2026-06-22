# MailTalk（仮）

OutlookのメールをLINE/Slack風のチャットUIで表示し、会話を**優先度の4レーン**に自動仕分けするローカルGUIアプリ。

> 設計仕様・思想の全文は [CLAUDE.md](CLAUDE.md) を参照（実装前に必読）。

## これは何を解決するか

敵は“量”ではなく**認知負荷**。「どれを開けばいいか迷う／自分への依頼が埋もれる／返したか分からない」を、**自分ボール判定（レーン仕分け）** 1つでまとめて解く。

| レーン | 意味 | 条件（優先順位カスケード） |
|--------|------|------|
| 🔴 自分ボール | あなたが返す番 | 最後が相手の発言 かつ 自分がToにいる |
| 🟠 活発 | やり取り進行中 | （🔴でない）かつ 直近の往復が多い |
| 🔵 共有(FYI) | 用件なし | （上2つでない）かつ 自分がCCのみ |
| ⚪ 静か | 動きなし | 残り全部 |

## 技術スタック

- Python 3.10+ / SQLite キャッシュ / FastAPI（`127.0.0.1` 限定）
- Outlook接続は **win32com（pywin32）のみ**。Microsoft Graph API は使わない（社外送信不可ポリシー）
- LLM不使用。レーン仕分けは100%ルールベース
- 依存最小のバニラHTML/CSS/JS フロント

## 実行環境（A端末 / B端末 / 開発機）

会社環境はOutlookとPythonが別端末に分かれているため、**アプリをexe化してA端末で動かす**（詳細は [CLAUDE.md §14](CLAUDE.md)）。

- **開発機（macOS等）**: `win32com` は動かない。`triage` / `aggregate` / `db` / `notify` などの純粋ロジックを **Fakeデータで完結検証**（`FakeOutlookSource`）。
- **B端末（Python・Outlook無し）**: PyInstallerで **exeをビルドする機械**。
- **A端末（Outlook・Python無し・exe起動可）**: ビルドした **exeを実行**。win32com＋Outlookが同一マシンに揃うので、リアルタイム取得も返信も動く。UIはA端末のブラウザで閲覧。

`source.get_default_source()` がOSを見て実装を自動選択（Windows→`Win32OutlookSource`、それ以外→Fake）するため、同じ上位コードがどこでも動く。

## セットアップ & 実行

```bash
# 依存インストール（開発機）
uv venv
uv pip install -e ".[dev]"

# コンソールで仕分け結果を確認（macOSはFakeデータ）
uv run mailtalk
uv run mailtalk --limit 200
uv run mailtalk --diagnostics   # 診断バンドルを logs/ へ

# サーバ起動（チャットUI: http://127.0.0.1:8765）
uv run python -m mailtalk.main

# テスト（macOSで実行可）
uv run pytest
```

### exeビルド（B端末）と実行（A端末）

```bash
# B端末（Windows・Python）でビルド
git pull
uv pip install -e ".[windows,build]"   # pywin32 + pyinstaller
pyinstaller mailtalk.spec               # → dist/MailTalk.exe

# dist/MailTalk.exe を共有フォルダ経由でA端末へコピー
# A端末で MailTalk.exe をダブルクリック
#   → 127.0.0.1:8765 が自動でブラウザに開く（Outlookはローカル＝同一マシン）
#   → DB(data/) とログ(logs/) は exe の隣に作られる
```

開発機で直接サーバを動かす場合（Fakeデータ）:
```bash
uv run python -m mailtalk.main
```

## モジュール構成

| ファイル | 役割 |
|----------|------|
| `models.py` | ドメインモデル（COM非依存） |
| `triage.py` | レーン仕分け（優先順位カスケード）★最重要 |
| `aggregate.py` | Message→Conversation 集約・triage信号計算 |
| `source.py` | 取得層インターフェース＋実装の自動選択 |
| `outlook_client.py` | win32com 実装（Windows実機専用） |
| `fake_outlook.py` | 開発機用Fake供給元 |
| `db.py` | SQLite キャッシュ |
| `notify.py` | 二層メッセージ／ログ基盤 |
| `sync.py` | 同期ワーカー＋ステートマシン |
| `reply.py` | 返信下書き生成（ReplyAll→Display、送信はしない） |
| `main.py` | FastAPI（エンドポイント・静的配信・ブラウザ自動起動） |
| `paths.py` | 実行形態（通常/exe）別のパス解決 |
| `static/index.html` | チャットUI |
| `run_app.py` / `mailtalk.spec` | exe化のエントリとPyInstaller設定（B端末でビルド） |

## 制約（職場ポリシー）

- メール内容を外部ネットワークへ送信しない（全ローカル完結）
- **メールを自動送信しない**。返信は下書きをOutlookで開くまで。送信は人間
- サーバは外部公開しない（`127.0.0.1` 固定）
