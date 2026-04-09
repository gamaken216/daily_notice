# Daily Notice System

PC起動時に **Google カレンダーの予定** と **Google ToDo リスト** をポップアップ表示する Python スクリプトです。

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![OS](https://img.shields.io/badge/OS-Windows%2010%2F11-blue)
![License](https://img.shields.io/badge/license-MIT-green)

---

## 📋 必要な環境

- Windows 10 / 11
- Python 3.10 以上
- Google アカウント
- インターネット接続

---

## 📁 ファイル構成

```
daily-notice/
├── daily_notice.py              # メインスクリプト
├── register_daily_notice.ps1    # タスクスケジューラー登録スクリプト
├── README.md                    # このファイル
├── credentials.json             # ★ 自分で用意（Google Cloud からダウンロード）
├── token.json                   # 初回認証後に自動生成
└── config.json                  # テーマ選択後に自動生成
```

> ⚠️ `credentials.json` / `token.json` / `config.json` は `.gitignore` で除外しています。  
> **絶対に Git にコミットしないでください。**

---

## 🚀 セットアップ手順

### Step 1: Google Cloud の設定

1. [Google Cloud Console](https://console.cloud.google.com/) にアクセスしてログイン
2. 新しいプロジェクトを作成
3. 以下の 2 つの API を有効化
   - **Google Calendar API**
   - **Tasks API**（「Cloud Tasks」ではなく「Tasks API」）
4. OAuth 同意画面 → 対象：「外部」/ テストユーザーに自分の Gmail を追加
5. 認証情報 → OAuth クライアント ID を作成（種類：デスクトップアプリ）
6. JSON をダウンロードし、`credentials.json` にリネームしてこのフォルダに配置

### Step 2: Python ライブラリのインストール

```bash
pip install google-auth google-auth-oauthlib google-api-python-client
```

### Step 3: 初回実行・Google 認証

```bash
python daily_notice.py
```

ブラウザが開いて Google 認証画面が表示されます。許可すると `token.json` が自動生成され、次回以降は認証不要になります。

### Step 4: PC 起動時の自動実行を登録

PowerShell を **管理者権限** で開き、このフォルダで以下を実行します：

```powershell
powershell -ExecutionPolicy Bypass -File "register_daily_notice.ps1"
```

次回 Windows ログオン時から自動でポップアップが表示されます。

---

## 🎨 テーマ（色）の変更

ポップアップ右上の「🎨 色変更」ボタンをクリックするだけで、8 種類のカラーテーマから選べます。選んだ設定は `config.json` に自動保存されます。

起動時にテーマ選択を先に表示したい場合：

```bash
python daily_notice.py --choose-color
```

| テーマ | 説明 |
|--------|------|
| 🌙 ダーク（デフォルト） | 暗い背景・青白い文字 |
| ☁️ クリーンホワイト | 白背景・黒文字 |
| 🌿 ナチュラルグリーン | 薄緑背景 |
| 🌅 ウォームベージュ | 温かみのあるベージュ |
| 🌌 ミッドナイトブルー | 深い紺色 |
| 🌸 ソフトピンク | 淡いピンク |
| 🌞 サンシャインイエロー | 明るいイエロー |
| 🪨 モノクロームグレー | シックなグレー |

---

## 🔧 トラブルシューティング

| 症状 | 対処法 |
|------|--------|
| `403: access_denied` | Google Cloud Console でテストユーザーに自分の Gmail を追加 |
| ポップアップが表示されない | タスクスケジューラーで「DailyNotice」の実行結果を確認 |
| 認証エラー・token 期限切れ | `token.json` を削除して再度 `python daily_notice.py` を実行 |
| Tasks API が見つからない | ライブラリで「Tasks API」を選択（「Cloud Tasks」は別サービス） |
| カレンダーの予定が出ない | primary カレンダーのみ対応。他のカレンダーはスクリプト修正が必要 |

---

## 🗑️ アンインストール

自動起動を解除する場合は PowerShell で：

```powershell
Unregister-ScheduledTask -TaskName "DailyNotice" -Confirm:$false
```

---

## 📄 ライセンス

MIT License
