# CLAUDE.md — MailTalk（仮）
OutlookのメールをLINE/Slack風のチャットUIで表示し、会話を「優先度の4レーン」に自動仕分けするローカルGUIアプリ。本ファイルは開発を引き継ぐClaude Code向けの設計仕様書。**実装前に必ず全体を読むこと。**
---
## 0. このプロジェクトの目的（最重要・ブレさせない）
ユーザー（日本の銀行・運用部署のエンジニア）の困りごとは「メールが読めない」ことではなく、**「読むべきものの選別と、文脈の再構築に脳のエネルギーを使わされる」**こと。敵は“量”ではなく“認知負荷”。
ユーザー自身が挙げた痛みの優先順位（上ほどツラい）:
1. **どれを開けばいいか迷う**（入口での迷い）← 最優先で解く
2. 開いても何の話か分からない（文脈の再構築）
3. 自分への依頼が埋もれる（用件の埋没）
4. 返したか分からない（取りこぼし）
5. 長い往復を読み直す（往復の重さ）
設計上の発見: **①③④は「これは自分ボールか？どれくらい急ぐか？」という同じ根の別の顔**。だから「自分ボール判定エンジン（＝レーン仕分け）」を1つ作れば①③④をまとめて潰せる。②⑤は「開いた後の読みやすさ」レイヤー。
**真の主役は「会話一覧画面（仕分けされたリスト）」**。チャット吹き出しは「開いた後」を楽にするもので、一覧は「開く前」を楽にするもの。①が最優先なので、力を注ぐのは一覧の仕分け精度。
MVPは「**読むのを楽にする**」一点に集中する。下記「将来機能」は土台が固まるまで実装しない。
---
## 1. 技術スタックと前提
- **OS**: Windows 専用（win32com / pywin32 で起動中のOutlookデスクトップを操作するため）
- **言語**: Python 3.x
- **Outlook接続**: `pywin32`（win32com）。**Microsoft Graph APIは使わない**（職場の「外部クラウドへのデータ送信不可」制約に抵触するため。すべてローカル完結）
- **キャッシュDB**: SQLite（永続化。後述の起動高速化の要）
- **バックエンド/サーバ**: FastAPI + uvicorn（ローカルのみbind: `127.0.0.1`）
- **フロント**: まずは依存の少ないバニラHTML/CSS/JSの単一ページをFastAPIから配信（社内プロキシ＋ホワイトリスト制約下でnpmビルド地獄を避けるため）。Reactは必要になってから検討
- **LLMは使わない（MVP）**: レーン仕分けは100%ルールベース。要約（痛み⑤）は将来機能で、使う場合も社内LLM APIのみ
### 制約・禁止事項（職場ポリシー由来）
- 外部ネットワークへメール内容を送信しない。テレメトリ・外部分析なし。全部ローカル
- **メールを自動送信しない**。返信は必ず「下書きをOutlookで開く」までで止め、最終送信は人間が押す（後述）
- サーバは外部公開しない（`127.0.0.1`固定）
### リポジトリ
- GitHub: `ShunsukeTamura06`（新規リポジトリを想定）
---
## 2. アーキテクチャ（3層）
```
GUI層 (チャットUI)        … 一覧（4レーン）/ 吹き出し / 検索 / 「トピックごと」「人ごと」切替
   ↑
変換・集約層              … Email→Message正規化 / Conversation・人でグルーピング / レーン仕分け
   ↑
取得層                    … win32comで起動中Outlookから取得 → SQLiteにキャッシュ
```
### 想定モジュール構成
- `outlook_client.py` — win32com接続、メール列挙、SMTP正規化、To/CC判別、会話組み立て
- `triage.py` — レーン仕分け（優先順位カスケード）
- `db.py` — SQLiteスキーマ・読み書き・差分同期状態
- `sync.py` — バックグラウンド同期ワーカー（新しい順・進捗）＋ステートマシン
- `notify.py` — 二層メッセージ基盤（ユーザー向け通知 / 開発者向けログ）
- `reply.py` — 返信下書き生成（ReplyAll → Display）
- `main.py` — FastAPIアプリ（エンドポイント・静的配信）
- `static/index.html` — チャットUI（後述モック準拠）
- `logs/app.log` — 開発者向けローテーションログ
---
## 3. レーン仕分け（最重要ロジック）
会話を点数で並べるのではなく、**上から順に条件を当て、最初に当たったレーンで確定する「優先順位カスケード」**。🔴を必ず最初に判定する。
| 順 | レーン | 条件 |
|----|--------|------|
| 1 | 🔴 自分ボール | `最後の発言が自分以外` かつ `自分がToに入っている` |
| 2 | 🟠 活発 | （🔴でない）かつ 直近24-48hの往復が多い |
| 3 | 🔵 共有(FYI) | （上2つでない）かつ 自分がCCのみ |
| 4 | ⚪ 静か | 残り全部 |
### 絶対に守る設計判断
- **🔴を活発より先に評価する。** 激活発スレッドでも「最後が相手＆自分がTo」なら問答無用で🔴。活発さは🔴を押しのけない。これにより「本当は自分ボールなのに活発に紛れて見逃す」が定義上起きない。
- **🔴のゲートは「既読/未読」ではなく「返信状態（最後が相手か）」。** 「読んだけど返してない」（痛み④）を取りこぼさないため。既読でも相手の発言で止まっていれば🔴に残す。未読はゲートではなく、UI上の強調表示にすぎない。
- **エラーコストは非対称→🔴は再現率優先。** 🔴の偽陽性（本当はFYIなのに🔴）はチラ見で閉じれば済む軽いコスト。🔴の偽陰性（見逃し）は重い。**To/CC判別が不能なとき等、迷ったら🔴に倒す。**
- **締め文の🔴誤検知はMVPでは許容。** 「ありがとうございました」等の返さなくていい相手発言が🔴に出ても、見れば分かるので潰さない。将来、③の依頼文検知や締め文検知で落とせる構造にしておく。
### 参考実装（pseudo）
```python
ACTIVE_THRESHOLD = 6  # 直近24-48hの件数しきい値。実データを見て調整可能にする
def classify(conv) -> tuple[str, str]:
    # 1. 🔴 自分ボール（最初に判定）
    if conv.last_from_other and conv.i_am_to:
        return "red", "Toに自分・最後が相手の発言（あなたが返す番）"
    # （将来）CC埋もれ依頼の救済:
    # if conv.last_from_other and conv.has_request_to_me:
    #     return "red", "CCだが本文であなたに依頼あり"
    # 2. 🟠 活発
    if conv.velocity_recent >= ACTIVE_THRESHOLD:
        return "amber", f"直近{conv.velocity_recent}件・活発にやり取り中"
    # 3. 🔵 共有(FYI)
    if conv.i_am_cc_only:
        return "blue", "CCのみ・あなた宛の用件なし"
    # 4. ⚪ 静か
    return "gray", "動きなし"
```
各会話には必ず `lane_reason`（人間が読める理由文）を持たせ、UIに表示する。「なぜ最優先なのか」が説明できることが、ルールベースの信頼性の核。
---
## 4. ビュー仕様
- **2ビュー切替**: 「トピックごと」（Outlookの`ConversationID`単位）と「人ごと」を行き来できる
- **CC多数メールはグループLINE風**: 参加者の顔ぶれを「グループ」として表示。グルーピングは `ConversationID` ベース＋参加者リストをグループ名/メンバーとして添える方式（参加者集合の完全一致や類似度クラスタリングは将来検討）
- **検索・フィルタ**: 件名・差出人で絞り込み
- **吹き出し表示**: 会話を時系列で、自分の発言は右、相手は左。差出人名・時刻つき。これが痛み②の解決
UIの見た目・情報設計は本リポジトリのモック（一覧＝レーン色ドット＋行タグ〔未読/To:自分/CCのみ/人数/件数〕、吹き出し先頭に🔴理由チップ）に準拠する。
---
## 5. 返信フロー（自動送信は絶対しない）
ユーザーがツールのチャット欄に本文を入力 →「Outlookで返信を開く」を押す → **署名入りの全員返信の下書きがOutlookで開く** → 人間が宛先・本文・署名を目視確認 → 人間が送信ボタンを押す。ツールは下書きを開くところまで。
- **常にReplyAll**（Outlookの「全員に返信」と完全に同じ挙動。元のTo＋CCから自分を除いた全員が宛先）。「大は小を兼ねる」＋職場文化。分岐ロジックは不要
- `.Send()` は**絶対に呼ばない**。必ず `.Display()` で作成ウィンドウを開くだけ
- 署名はOutlook自動挿入。COMの`.Reply()`系だと入らないことがあるので、`.GetInspector` にアクセスして作成画面を初期化させると署名＋引用元が `HTMLBody` に入る。その先頭にユーザー本文を差し込む（→「本文 → 署名 → 引用元」の標準レイアウト）
- 返信対象は会話内の**最新メール**
### 参考実装
```python
def open_reply_draft(ns, entry_id, store_id, body_text):
    original = ns.GetItemFromID(entry_id, store_id)   # 会話内の最新メール
    reply = original.ReplyAll()                        # 常に全員返信
    _ = reply.GetInspector          # 署名を HTMLBody に挿入させる
    base = reply.HTMLBody           # = 署名 + 引用元スレッド
    typed = body_text.replace("\n", "<br>")
    reply.HTMLBody = f"<div>{typed}</div>{base}"
    reply.Display()                 # 作成ウィンドウを開くだけ。送信は人間
```
---
## 6. 起動・準備完了モデル
理想は「ユーザーがPCとOutlookと本アプリを起動 → 裏で初期準備が進み → 開く頃には使える状態」。
- **永続SQLiteキャッシュ**: 初回（コールド）だけ全件読み込みで時間がかかる。2日目以降は前回以降の**差分だけ取り込む（インクリメンタル同期）**ので一瞬で温まる
- **新しい順に同期**: 直近メール（＝🔴になりやすい今朝の対応分）を先に読み込み、古い過去ログは裏で後追い補完
- **進捗の可視化（プログレッシブ）**: 準備中にユーザーが開いても、空画面ではなく〔進捗（例「メールを読み込み中… 1,240 / 3,500件」）＋すでに読み込めた会話〕を見せる。**直近は使える、古い分だけ読み込み中**という止まらない体験
- **Outlook起動待ち**: 本アプリがOutlookより先に立ち上がっても壊れないよう、接続をリトライしながら待機し、起動を検知したら自動で進む
### ステートマシン
`起動中 → Outlook接続中 → 同期中(◯/◯件) → 仕分け中 → 準備完了`、異常時は `エラー(原因つき)`。UIがいつでも現在状態と進捗を問い合わせられるよう、状態をエンドポイント（またはSSE）で公開する。
---
## 7. 二層メッセージ／ログ（最初から作り込む）
**1か所の呼び出しで行き先を2つに分ける。**
- **ユーザー向け**: 平易な日本語の状態・警告をUIに表示。例「Outlookが起動していません。起動すると読み込みを始めます」「準備完了。3,500件を仕分けしました」
- **開発者向け**: 詳細ログをローテーション付きでファイル（`logs/app.log`）に残す。COMエラーコード、DN→SMTP変換の失敗、各処理の所要時間などを全部
`notify_user(level, message)` を呼ぶと、UI通知（平易）と詳細ログ（技術情報）の両方へ適切な粒度で流れる設計にする。例: Outlook未起動なら、ユーザーには一文、ログには `com_error -2147221021 / CoCreateInstance Outlook.Application` のような生情報。
---
## 8. データモデル（SQLite・たたき台）
```sql
CREATE TABLE emails (
  entry_id         TEXT PRIMARY KEY,   -- Outlook EntryID
  store_id         TEXT,               -- GetItemFromID に entry_id と併せて必要
  conversation_id  TEXT,               -- Outlook ConversationID
  subject          TEXT,
  subject_norm     TEXT,               -- Re:/Fwd: 除去後
  sender_email     TEXT,               -- 正規化済みSMTP
  sender_name      TEXT,
  to_json          TEXT,               -- SMTPのJSON配列
  cc_json          TEXT,
  is_to_me         INTEGER,            -- 0/1
  is_cc_me         INTEGER,
  is_from_me       INTEGER,
  body_preview     TEXT,
  body_html        TEXT,               -- 吹き出し表示用（サイズ次第で遅延取得も可）
  received_time    TEXT,               -- ISO8601
  unread           INTEGER,
  importance       INTEGER,
  folder           TEXT,
  synced_at        TEXT
);
CREATE TABLE conversations (
  conversation_id   TEXT PRIMARY KEY,
  subject_norm      TEXT,
  participants_json TEXT,
  participant_count INTEGER,
  last_received     TEXT,
  last_sender_email TEXT,
  last_from_me      INTEGER,
  any_unread        INTEGER,
  i_am_to           INTEGER,           -- 直近/いずれかでToに自分
  i_am_cc_only      INTEGER,
  velocity_recent   INTEGER,           -- 直近24-48hの件数
  lane              TEXT,              -- red / amber / blue / gray
  lane_reason       TEXT,
  updated_at        TEXT
);
CREATE TABLE contacts (
  email        TEXT PRIMARY KEY,
  display_name TEXT
);
CREATE TABLE sync_state (
  key   TEXT PRIMARY KEY,             -- last_sync_time, status, progress 等
  value TEXT
);
```
`lane`/`lane_reason` は同期後の仕分けで書き込む（オンザフライ計算でも可だが、UI高速化のためキャッシュ推奨）。
---
## 9. win32comのハマりどころ（精度はここで決まる）
レーン仕分けロジック自体は単純で、🔴が正確かは**入力信号の精度**でほぼ決まる。テストはここを重点的に。
### (a) 「自分」を正しく特定する
`SenderEmailAddress` はExchange環境だとSMTPでなく `/o=ExchangeLabs/...` 形式のDNが返ることがある。単純比較だと外す。
```python
def my_smtp_addresses(ns):
    addrs = set()
    try:
        exu = ns.CurrentUser.AddressEntry.GetExchangeUser()
        if exu:
            addrs.add(exu.PrimarySmtpAddress.lower())
    except Exception:
        pass
    for acc in ns.Session.Accounts:
        if getattr(acc, "SmtpAddress", None):
            addrs.add(acc.SmtpAddress.lower())
    return addrs
def sender_smtp(mail):
    if getattr(mail, "SenderEmailType", "") == "EX":
        try:
            return mail.Sender.GetExchangeUser().PrimarySmtpAddress.lower()
        except Exception:
            return (mail.SenderEmailAddress or "").lower()
    return (mail.SenderEmailAddress or "").lower()
```
### (b) To と CC を正しく分ける
`Recipients` の `Type`（1=To, 2=CC, 3=BCC）で判定。EX→SMTP変換も同様に必要。配布リスト宛だと自分が直接Toに見えないことがある→**判別不能なときは安全側＝🔴寄りに倒す**。
```python
# r.Type: 1=To, 2=CC, 3=BCC
for r in mail.Recipients:
    smtp = resolve_smtp(r.AddressEntry)  # EX なら GetExchangeUser().PrimarySmtpAddress
    if r.Type == 1: to_list.append(smtp)
    elif r.Type == 2: cc_list.append(smtp)
```
### (c) 「最後の発言」を正しく取る
会話内アイテムを `ReceivedTime` で並べて末尾を取る。自動の開封通知や締め文も「最後が相手」になり🔴誤検知を生むが、MVPでは許容（§3参照）。
### (d) その他
- COM呼び出しは例外（`pywintypes.com_error`）を握って二層ログへ。1通の失敗で全体を止めない
- 大量メールの列挙は重い。フォルダの `Items.Restrict` / `Sort` で対象期間・新しい順に絞ってから回す
- スレッド/プロセスをまたぐCOM利用は `pythoncom.CoInitialize()` に注意（バックグラウンド同期ワーカー）
---
## 10. 開発の進め方（推奨ビルド順）
土台がズレるとガワを作り直す羽目になるので、**土台→ガワ**の順で。
1. **コア先行**: `outlook_client` + `triage` + `db` + 二層ログ基盤を作り、**「会話一覧を読み込んで、各会話がどのレーンか／最後が誰の発言か／自分はTo・CCどっちか、をコンソールに吐く」**ところまで。実データで仕分け精度を確認する
2. 精度が信頼できたら `sync`（新しい順・進捗・ステートマシン・差分同期）を追加
3. `main.py`（FastAPI）でエンドポイント化（一覧をレーン別取得 / 会話の吹き出し取得 / 状態・進捗取得 / 返信下書きを開く）
4. `static/index.html` でUIをモック準拠に実装
5. `reply.py`（ReplyAll→Display）を組み込み、返信フローを通す
各段階でユーザーがWindows実機で動作確認する（win32comはWindows+Outlook起動が必須で、Windows以外では実行・テスト不可）。
---
## 11. 将来機能（土台が固まるまで実装しない）
MVPで「読むのが楽」を達成・体感確認してから、上に積む:
- スレッド3行要約（社内LLM API。外部送信不可なので社内APIのみ）= 痛み⑤
- 自分宛の依頼文ハイライト＝痛み③（§3のCC埋もれ救済で🔴へ昇格も）
- 未返信レーダー、人物カード、添付ファイル一覧、ピン留め、対応済✅などの自分用ステータス
- 会話のトピックネットワーク化、意思決定ログ自動抽出
- 事務担当者向けの「設定不要・起動したら勝手に表示」超シンプルモード（将来のチーム展開を見据える）
---
## 12. 用語
- **自分ボール**: 相手の発言で止まっていて、自分が返す番の会話（🔴）
- **レーン**: 会話一覧の仕分け区分（🔴自分ボール / 🟠活発 / 🔵共有FYI / ⚪静か）
- **会話 / Conversation**: Outlookの `ConversationID` で束ねたスレッド

---
## 13. 開発メモ（実装側で追記）

### 環境の二重性（重要）
- **開発機**: macOS。`win32com` は動かない（Windows専用）。よってここでは **triage / db / models / notify / グルーピング等の純粋ロジックを Fake データで完結検証**する。
- **実機**: Windows + Outlook 起動必須。`outlook_client.py` の COM 依存部分はここでしか実行・テストできない。
- 設計方針: COM 依存を `outlook_client.py` に閉じ込め、上位層は `models.Message` / `models.Conversation` のみに依存させる。`fake_outlook.py` が同じインターフェースを提供し、macOS でパイプライン全体（取得→グルーピング→仕分け→DB→表示）を流せる。

### モジュール配置
- `src/mailtalk/` 配下にパッケージ。`uv run` で実行。テストは `tests/`（pytest、macOSで実行可）。

### 持ち帰り最小化（会社端末←→開発機）
- 会社端末でしか出ない COM 挙動の調査は `python -m mailtalk.cli --diagnostics` 一発で〔仕分け結果・所要時間・DN→SMTP変換失敗・環境情報〕を `logs/` にまとめる方針（シークレットは出さない）。

### レビュー反映済みの精度仕様（🔴の正確さに直結・変更不可）
codexレビューを受けて以下を確定。triageのカスケード自体は §3 のままだが、入力信号の作り方を厳密化した:
- **送信済みも取得する**: `outlook_client` は受信トレイ＋送信済み(olFolderSentMail=5)を読む。これがないと返信済みでも `last_from_me=False` になり🔴に残る（痛み④を取りこぼす）。
- **宛先信号は「最新の相手メール」基準**: `aggregate` は会話の最後の相手発言(=いま返す相手)を基準に `i_am_to`/`i_am_cc_only` を計算する。過去にToだっただけで🔴にしない。
- **解決不能Toは🔴寄り（§9準拠）**: 配布リスト宛やEX→SMTP変換失敗のToは `Message.to_unresolved=True` とし、`i_am_to` 扱いに倒す。EX解決は GetExchangeUser→PropertyAccessor(PR_SMTP_ADDRESS)→Address の順でフォールバック。
- **自分アドレス空は明示失敗**: `my_addresses()` が空なら黙って継続せずエラー停止（is_from_me/is_to_me全滅による🔴全滅を防ぐ）。
- **列挙の例外で全体を止めない**: GetFirst/GetNext を含めフォルダ列挙を例外保護（§9d）。
- **未対応（許容・将来）**: `velocity_recent` は単純件数で「往復(送信者交代)」を見ていない。🟠は🔴の後段で最低リスクのためMVP許容、実データで閾値/指標を調整する。

---
## 14. 配布・デプロイ（確定：exe on A端末・単一マシン）

### 端末構成（実環境の制約）
会社には2台あり、当初の「Python＋Outlook同一マシン」前提が崩れていた:
- **A端末**: Outlookあり。**Pythonインタプリタは使えない**が、**未署名exeは起動可能**（確認済み）。**UIはこのA端末のブラウザで見る**。
- **B端末**: Pythonが使える別端末。**Outlookは無い**。
- A・Bはネットワーク接続。共有フォルダで読み書き可能。

### 確定した解
**アプリ全体をPyInstallerで単一exeに固め、A端末で実行する。** これにより
Python・pywin32(win32com)・FastAPI/uvicorn・staticがexe内に揃い、Outlookと
同一マシン(A端末)で動くため **win32comがそのまま使え、リアルタイム取得も返信
(ReplyAll→Display)も成立する**。`source.get_default_source()` はWindowsで
`Win32OutlookSource` を選ぶので、追加の分岐は不要。

> 検討して**不採用**にした案: 共有フォルダ経由のJSON受け渡し（A端末でPS抽出→B端末で消費）。リアルタイム性が無く返信もできないため却下。PowerShell常駐HTTP案は、exeが起動可能と判明したため保留（exeが将来弾かれた場合の代替）。

### 役割分担とフロー
```
このPC(mac) : 開発・純ロジック検証(Fake) → GitHub push
B端末       : git pull → PyInstallerでビルド（mac不可なのでビルド機はB端末）
              uv pip install -e ".[windows,build]" ; pyinstaller mailtalk.spec
共有フォルダ : dist/MailTalk.exe をA端末へ受け渡し
A端末       : MailTalk.exe をダブルクリック → 127.0.0.1:8765 が自動でブラウザに開く
              → 閲覧・返信（Outlookはローカル＝同一マシン）
```

### exe化の要点（`mailtalk.spec` / `run_app.py` / `paths.py`）
- **エントリ**: `run_app.py`（`multiprocessing.freeze_support()` 後に `main()`）。
- **同梱**: `static/index.html` を datas で同梱。win32com/uvicorn/pydanticは動的importが多いので `collect_submodules` で明示収集。
- **パス解決**: `paths.py` が frozen を判定し、読み取り専用リソース(static)は `sys._MEIPASS`、書き込み(data/logs)は **exeと同じフォルダ** に置く（exeの隣にDB・ログが溜まる＝診断・持ち帰りが容易）。
- **起動UX**: `main()` が起動後に既定ブラウザで `127.0.0.1:8765` を自動オープン。`console=True` で接続状況・エラーをその場で確認可能。
- **返信**: A端末ではローカルCOMなので §5 の `open_reply_draft`（ReplyAll→Display、`.Send()`禁止）がそのまま機能する。
