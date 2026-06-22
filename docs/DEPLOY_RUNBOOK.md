# 実機デプロイ手順書（S端末ビルド → M端末実行）

MailTalk を会社端末で動かすための手順とトラブルシュート。**S端末でビルド → 共有フォルダ → M端末で実行**。
迷ったらこの順に進める。問題が出たら最後の「§6 失敗時」で診断zipを取って持ち帰る（1回で済ます）。

- **S端末（開発端末）**: Python可・`git clone`可・PyPI/社内ミラー接続可。**ビルド機**。
- **M端末（社用端末）**: Outlookあり・Python不可・未署名exe起動可。**実行機**（UIもここのブラウザで見る）。
- 前提: S↔M はネットワーク接続、共有フォルダで読み書き可。

---

## 1. S端末: ビルド

```powershell
# 初回のみ
git clone https://github.com/ShunsukeTamura06/mail-talk.git
cd mail-talk
# 2回目以降は最新を取得
git pull

# 依存インストール（pywin32 + pyinstaller を含む）
uv venv
uv pip install -e ".[windows,build]"

# 単一exeをビルド
uv run pyinstaller mailtalk.spec
# → dist\MailTalk.exe が生成される
```

uv が無い場合（標準のvenv/pip）:
```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -e ".[windows,build]"
pyinstaller mailtalk.spec
```

### ビルド直後の動作確認（S端末でできる範囲）
S端末にOutlookは無いので「Outlook未接続」になるのが**正常**。確認したいのは**exeが起動し、UIが配信されるか**だけ:
```powershell
dist\MailTalk.exe
```
- 黒いコンソール窓が開く → 数秒で既定ブラウザが `http://127.0.0.1:8765` を開く
- UIが表示され、状態が「Outlook接続中…」→（最大約60秒後）「読み込みに失敗（Outlook未起動）」になればビルドはOK
- コンソール窓を閉じれば終了

---

## 2. ビルドのトラブルシュート

| 症状 | 原因 | 対処 |
|---|---|---|
| 実行時 `No module named 'win32timezone'` 等 | pywin32の動的import取りこぼし | `mailtalk.spec` で収集済み。出る場合は `pip install --force-reinstall pywin32` 後、`python .venv\Scripts\pywin32_postinstall.py -install` を実行して再ビルド |
| 実行時 `DLL load failed`（pythoncom/pywintypes） | pywin32のDLL未同梱 | pywin32を正規手順で入れ直す（上記postinstall）。それでも出れば spec に `--collect-binaries pywin32` 相当を追加 |
| `uvicorn`/`click` 等の `ModuleNotFoundError` | submodule取りこぼし | spec の `collect_submodules` に該当パッケージを追加して再ビルド |
| 起動後 `index.html` が無い旨のエラー | staticが未同梱 | **リポジトリのルートで**ビルドする（spec の `Tree("src/mailtalk/static", ...)` がルート相対）。`dist`/`build` を消して再ビルド |
| ビルド物がアンチウイルスに隔離される | PyInstallerブートローダの誤検知 | `dist` フォルダ/exe をAV例外に追加。社内ポリシー次第でIT申請 |

困ったら S端末で `uv run pyinstaller mailtalk.spec --clean` で作り直す。

---

## 3. 受け渡し（共有フォルダ）

1. S端末の `dist\MailTalk.exe` を共有フォルダへコピー
2. M端末で**ローカルへコピー**（例: `C:\Users\<自分>\MailTalk\MailTalk.exe`）
   - 共有フォルダ上から直接起動でも動く（書込先は `%LOCALAPPDATA%\MailTalk` に自動フォールバック）が、**ローカルにコピーして起動を推奨**（速度・書込みの確実性）

---

## 4. M端末: 実行

1. **Outlookを起動**しておく（先でも後でもよい。アプリは接続をリトライして待つ）
2. `MailTalk.exe` をダブルクリック
3. 黒いコンソール窓が開く → 数秒で既定ブラウザが `http://127.0.0.1:8765` を開く
4. コンソールに「Outlookに接続しました」、UIヘッダが「準備完了。N件の会話を仕分けしました」になれば成功
5. 🔴自分ボールから確認していく。返信は会話を開いて本文入力 →「Outlookで返信を開く」→ **下書きが開くので人間が送信**（ツールは送信しない）

終了: コンソール窓を閉じる。

---

## 5. 実行のトラブルシュート

| 症状 | 対処 |
|---|---|
| 「WindowsによってPCが保護されました」(SmartScreen) | 「詳細情報」→「実行」。未署名exeのため毎回出ることがある |
| AppLocker等でブロック | 許可されたフォルダから起動、またはIT例外申請 |
| いつまでも「Outlook接続中…」 | Outlookが未起動/応答なし。Outlookを起動・再起動すると自動で進む（最大約60秒待つ） |
| `com_error -2147221021` 等がコンソールに | Outlookが応答していない。Outlook再起動。直らなければ §6 で診断zip取得 |
| ブラウザが自動で開かない | 手動で `http://127.0.0.1:8765` を開く |
| `Address already in use`（8765） | 既に起動済み。多重起動なら一方を終了。残プロセスはタスクマネージャで終了 |
| 一覧は出るが🔴がおかしい | §6 で診断zipを取得して持ち帰る（精度はそのzipで判断） |

---

## 6. 失敗時: 診断zipを取って持ち帰る（最重要・1回で済ます）

何か変なら、**原因究明に必要な情報を1つのzipにまとめて持ち帰る**。メール本文は含めない。

取得方法は2つ（どちらでも可）:

- **UIのボタン**: ヘッダ右の「診断zip」をクリック → 保存先パスがダイアログに出る
- **コマンド**: コマンドプロンプトで
  ```
  MailTalk.exe --diagnostics
  MailTalk.exe --diagnostics --redact   # アドレス・件名もマスク
  ```

出力: exe隣（書込不可なら `%LOCALAPPDATA%\MailTalk`）の `logs\mailtalk_diag_<日時>.zip`。
**このzip＋（あれば）コンソール窓の文言**を申請して持ち帰る。

### zipの中身（diagnostics.json）で分かること
- `errors`: 取得・接続・自己特定で起きた失敗（最優先で見る）
- `self_identification.my_addresses`: 自分のアドレスを正しく特定できたか（空なら🔴全滅の原因）
- `summary.lane_counts`: レーン内訳（🔴が0/多すぎ等の異常検知）
- `summary.to_unresolved_messages`: To解決不能（配布リスト等）の件数
- `conversations[].messages[]`: 各メールの `is_from_me/is_to_me/is_cc_me/to_unresolved/folder` と最後が誰か → 🔴判定の根拠を1通単位で検証できる
- `timing_sec`: 取得・仕分けの所要時間

---

## 7. 精度を見るときの観点（CLAUDE.md §3/§9）
- 🔴は「最後が相手 かつ 自分がTo（解決不能Toは🔴寄り）」。**返信済みなら🔴から外れる**＝送信済みフォルダの取得が効いているか（`folder` に Sent 系が出るか）を確認
- 自分宛か判別不能（配布リスト等）が多い場合は `to_unresolved` 件数を確認
- 閾値 `ACTIVE_THRESHOLD`（🟠）は実データで調整可（`triage.py`）
