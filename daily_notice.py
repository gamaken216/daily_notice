"""
PC起動時に今日のGoogleカレンダー予定とToDoリストをポップアップ表示するスクリプト
"""

import os
import sys
import json
import socket
import datetime
import tkinter as tk
from tkinter import scrolledtext
import urllib.request
import urllib.parse
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import base64
from email.mime.text import MIMEText

SCOPES = [
    "https://www.googleapis.com/auth/calendar.readonly",
    "https://www.googleapis.com/auth/tasks.readonly",
    "https://www.googleapis.com/auth/gmail.send",
]

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_FILE = os.path.join(SCRIPT_DIR, "credentials.json")
HOSTNAME = socket.gethostname()
TOKEN_FILE = os.path.join(SCRIPT_DIR, f"token_{HOSTNAME}.json")
CONFIG_FILE = os.path.join(SCRIPT_DIR, f"config_{HOSTNAME}.json")
LAST_SENT_FILE = os.path.join(SCRIPT_DIR, "last_sent.json")  # 全PC共通: 最終送信日

# ============================================================
# 🎨 カラーパレット定義
# ============================================================
COLOR_PALETTES = [
    {"name": "🌙 ダーク（デフォルト）",    "bg": "#1e1e2e", "fg": "#cdd6f4", "title": "#89b4fa", "btn_bg": "#89b4fa", "btn_fg": "#1e1e2e"},
    {"name": "☁️ クリーンホワイト",        "bg": "#ffffff", "fg": "#1a1a1a", "title": "#2d6a4f", "btn_bg": "#2d6a4f", "btn_fg": "#ffffff"},
    {"name": "🌿 ナチュラルグリーン",       "bg": "#e8f5e9", "fg": "#1b5e20", "title": "#388e3c", "btn_bg": "#388e3c", "btn_fg": "#ffffff"},
    {"name": "🌅 ウォームベージュ",         "bg": "#fdf6ec", "fg": "#3e2723", "title": "#bf360c", "btn_bg": "#bf360c", "btn_fg": "#ffffff"},
    {"name": "🌌 ミッドナイトブルー",       "bg": "#0d1b2a", "fg": "#e0e0e0", "title": "#90caf9", "btn_bg": "#1565c0", "btn_fg": "#ffffff"},
    {"name": "🌸 ソフトピンク",             "bg": "#fce4ec", "fg": "#4a0030", "title": "#c2185b", "btn_bg": "#c2185b", "btn_fg": "#ffffff"},
    {"name": "🌞 サンシャインイエロー",     "bg": "#fffde7", "fg": "#3e2700", "title": "#f57f17", "btn_bg": "#f57f17", "btn_fg": "#ffffff"},
    {"name": "🪨 モノクロームグレー",       "bg": "#2b2b2b", "fg": "#e0e0e0", "title": "#bdbdbd", "btn_bg": "#616161", "btn_fg": "#ffffff"},
]
# ============================================================


def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"palette_index": 0}


def save_config(config):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


def choose_color_palette(parent, current_index=0):
    """
    色選択ダイアログ。parent ウィンドウの前面に表示する。
    選択されたインデックスを返す。キャンセル時は current_index を返す。
    """
    selected_index = [current_index]

    dialog = tk.Toplevel(parent)
    dialog.title("色テーマを選んでください")
    dialog.geometry("400x500")
    dialog.resizable(False, False)
    dialog.configure(bg="#f5f5f5")
    dialog.transient(parent)       # 親ウィンドウに紐付け
    dialog.grab_set()              # このダイアログが閉じるまで親を操作不可に
    dialog.lift()                  # 最前面に表示
    dialog.attributes("-topmost", True)

    tk.Label(dialog, text="色テーマを選んでください",
             font=("Meiryo UI", 13, "bold"), bg="#f5f5f5", fg="#1a1a1a", pady=14).pack()

    frame = tk.Frame(dialog, bg="#f5f5f5")
    frame.pack(fill="both", expand=True, padx=20)

    var = tk.IntVar(value=current_index)

    # プレビューエリア（先に作る）
    preview_frame = tk.Frame(dialog, bg="#f5f5f5")
    preview_frame.pack(fill="x", padx=20, pady=(4, 2))
    p0 = COLOR_PALETTES[current_index]
    preview_title = tk.Label(preview_frame, text="2026年04月10日(Fri)  ← プレビュー",
                             font=("Meiryo UI", 10, "bold"), bg=p0["bg"], fg=p0["title"],
                             padx=10, pady=4, anchor="w")
    preview_title.pack(fill="x")
    preview_body = tk.Label(preview_frame, text=f'背景: {p0["bg"]}  文字: {p0["fg"]}',
                            font=("Meiryo UI", 10), bg=p0["bg"], fg=p0["fg"],
                            padx=10, pady=4, anchor="w")
    preview_body.pack(fill="x")

    def update_preview(idx):
        p = COLOR_PALETTES[idx]
        preview_title.configure(bg=p["bg"], fg=p["title"])
        preview_body.configure(bg=p["bg"], fg=p["fg"], text=f'背景: {p["bg"]}  文字: {p["fg"]}')

    for i, palette in enumerate(COLOR_PALETTES):
        tk.Radiobutton(frame, text=palette["name"], variable=var, value=i,
                       font=("Meiryo UI", 11), bg="#f5f5f5", activebackground="#e0e0e0",
                       command=lambda idx=i: update_preview(idx), anchor="w"
                       ).pack(fill="x", pady=1)

    def on_ok():
        selected_index[0] = var.get()
        dialog.destroy()

    btn_frame = tk.Frame(dialog, bg="#f5f5f5")
    btn_frame.pack(pady=10)
    tk.Button(btn_frame, text="この色に決定", command=on_ok,
              font=("Meiryo UI", 11, "bold"), bg="#2d6a4f", fg="#ffffff",
              relief="flat", padx=16, pady=6).pack(side="left", padx=8)
    tk.Button(btn_frame, text="キャンセル", command=dialog.destroy,
              font=("Meiryo UI", 11), bg="#9e9e9e", fg="#ffffff",
              relief="flat", padx=16, pady=6).pack(side="left", padx=8)

    parent.wait_window(dialog)     # ダイアログが閉じるまで待機
    return selected_index[0]


def get_credentials():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as token:
            token.write(creds.to_json())
    return creds


def get_calendar_events(creds):
    service = build("calendar", "v3", credentials=creds)
    now = datetime.datetime.now()
    start_of_day = now.replace(hour=0, minute=0, second=0, microsecond=0).isoformat() + "+09:00"
    end_of_day   = now.replace(hour=23, minute=59, second=59, microsecond=0).isoformat() + "+09:00"
    result = service.events().list(
        calendarId="primary", timeMin=start_of_day, timeMax=end_of_day,
        singleEvents=True, orderBy="startTime",
    ).execute()
    return result.get("items", [])


def get_tasks(creds):
    service = build("tasks", "v1", credentials=creds)
    tasklists = service.tasklists().list().execute().get("items", [])
    all_tasks = []
    for tl in tasklists:
        tasks = service.tasks().list(
            tasklist=tl["id"], showCompleted=False, showHidden=False
        ).execute().get("items", [])
        for t in tasks:
            t["_list_name"] = tl["title"]
        all_tasks.extend(tasks)
    return all_tasks


def format_time(event):
    start = event["start"].get("dateTime", event["start"].get("date", ""))
    if "T" in start:
        return datetime.datetime.fromisoformat(start).strftime("%H:%M")
    return "終日"


def get_airtable_followups(config):
    """Airtableの版権管理テーブルからフォローアップの案件を取得する"""
    token = config.get("airtable_token", "")
    base_id = config.get("airtable_base_id", "")
    table_id = config.get("airtable_table_id", "")
    if not token or not base_id or not table_id:
        return {}

    # フィールドID
    fld_book = "fld9s7dffl6y45Ivh"          # 書籍名（multipleLookupValues）
    fld_publisher = "fldMD5l4sCptEIFyQ"     # 海外出版社またはエージェント（singleLineText）
    fld_lang = "fld63md3EQb2jDzwf"          # 言語（multipleSelects）
    fld_status = "fldfy0QSYNGn3bcDu"        # ステータス（singleSelect）

    # 抽出条件:
    # 1. ステータス="オファー" → 社内判断待ち
    # 2. ステータス="オファー受諾" → 契約書未締結
    # 3. ステータス="契約書締結" AND 受取前払金金額=BLANK AND 前払金入金日=BLANK → 未入金
    formula = (
        'OR('
        '{ステータス}="オファー",'
        '{ステータス}="オファー受諾",'
        'AND({ステータス}="契約書締結",{受取前払金金額}=BLANK(),{前払金入金日}=BLANK())'
        ')'
    )

    params = urllib.parse.urlencode({
        "filterByFormula": formula,
        "fields[]": [fld_book, fld_publisher, fld_lang, fld_status],
    }, doseq=True)

    url = f"https://api.airtable.com/v0/{base_id}/{table_id}?{params}"
    req = urllib.request.Request(url, headers={
        "Authorization": f"Bearer {token}",
    })

    # カテゴリ別に分類
    categories = {
        "社内判断待ち": [],
        "契約書未締結": [],
        "未入金":       [],
    }
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        for record in data.get("records", []):
            fields = record.get("fields", {})
            book_raw = fields.get("書籍名", [])
            book = str(book_raw[0]) if isinstance(book_raw, list) and book_raw else "(不明)"
            publisher = fields.get("海外出版社またはエージェント", "")
            lang_raw = fields.get("言語", [])
            lang = ", ".join(item["name"] if isinstance(item, dict) else str(item) for item in lang_raw) if lang_raw else ""
            status = fields.get("ステータス", "")
            if status == "オファー":
                categories["社内判断待ち"].append((book, publisher, lang))
            elif status == "オファー受諾":
                categories["契約書未締結"].append((book, publisher, lang))
            elif status == "契約書締結":
                categories["未入金"].append((book, publisher, lang))
    except Exception as e:
        print(f"Airtable取得エラー: {e}", file=sys.stderr)

    return categories


def get_demo_events():
    today = datetime.date.today().isoformat()
    return [
        {"summary": "朝のミーティング",   "start": {"dateTime": f"{today}T09:00:00+09:00"}},
        {"summary": "企画書レビュー",     "start": {"dateTime": f"{today}T11:00:00+09:00"}},
        {"summary": "ランチ（社外）",     "start": {"dateTime": f"{today}T12:30:00+09:00"}},
        {"summary": "週次定例",           "start": {"dateTime": f"{today}T15:00:00+09:00"}},
        {"summary": "資料作成",           "start": {"date": today}},
    ]


def get_demo_tasks():
    return [
        {"title": "四半期レポートを提出する",       "_list_name": "仕事",         "due": ""},
        {"title": "新規プロジェクトの提案書を作成", "_list_name": "仕事",         "due": ""},
        {"title": "歯医者の予約を入れる",           "_list_name": "プライベート", "due": ""},
        {"title": "読みかけの本を読み終える",       "_list_name": "プライベート", "due": ""},
    ]


def show_popup(events, tasks, palette_index, followups=None):
    """メインポップアップ。palette_index を保持し、色変更時はウィジェットを再描画する。"""
    if followups is None:
        followups = {}
    today = datetime.date.today().strftime("%Y年%m月%d日(%a)")
    title_font   = ("Meiryo UI", 14, "bold")
    section_font = ("Meiryo UI", 12, "bold")
    body_font    = ("Meiryo UI", 11)

    root = tk.Tk()
    root.title(f"Today's Notice - {today}")
    root.geometry("520x600")
    root.attributes("-topmost", True)

    def build_ui(p):
        """全ウィジェットを破棄して新しいパレットで再構築する。"""
        for w in root.winfo_children():
            w.destroy()
        root.configure(bg=p["bg"])

        # ヘッダー
        header = tk.Frame(root, bg=p["bg"])
        header.pack(fill="x")
        tk.Label(header, text=today, font=title_font,
                 bg=p["bg"], fg=p["fg"], pady=10, anchor="w", padx=15
                 ).pack(side="left", fill="x", expand=True)

        def on_change_color():
            config = load_config()
            new_idx = choose_color_palette(root, current_index=config.get("palette_index", 0))
            config["palette_index"] = new_idx
            save_config(config)
            build_ui(COLOR_PALETTES[new_idx])   # ウィジェットを再構築

        tk.Button(header, text="🎨 色変更", command=on_change_color,
                  font=("Meiryo UI", 9), bg=p["btn_bg"], fg=p["btn_fg"],
                  relief="flat", padx=8, pady=4
                  ).pack(side="right", padx=10, pady=8)

        # テキストエリア
        text = scrolledtext.ScrolledText(
            root, font=body_font, bg=p["bg"], fg=p["fg"],
            insertbackground=p["fg"], wrap="word", relief="flat", padx=15, pady=10)
        text.pack(fill="both", expand=True)

        text.insert("end", "--- Google Calendar ---\n\n", "section")
        if events:
            for event in events:
                text.insert("end", f"  {format_time(event)}  {event.get('summary','(無題)')}\n")
        else:
            text.insert("end", "  予定はありません\n")
        text.insert("end", "\n")

        text.insert("end", "--- ToDo List ---\n\n", "section")
        if tasks:
            current_list = ""
            for task in tasks:
                ln = task.get("_list_name", "")
                if ln != current_list:
                    text.insert("end", f"  [{ln}]\n")
                    current_list = ln
                due_str = ""
                if task.get("due"):
                    d = datetime.datetime.fromisoformat(task["due"].replace("Z", "+00:00"))
                    due_str = f" (期限: {d.strftime('%m/%d')})"
                text.insert("end", f"    - {task.get('title','(無題)')}{due_str}\n")
        else:
            text.insert("end", "  タスクはありません\n")

        # Ascom 要フォローアップセクション（該当ありの場合のみ）
        has_followups = any(items for items in followups.values())
        if has_followups:
            text.insert("end", "\n")
            text.insert("end", "--- ⚠️ Ascom 要フォローアップ ---\n\n", "section_warn")
            category_labels = {
                "社内判断待ち": "【社内判断待ち】オファーが出たまま社内判断保留中",
                "契約書未締結": "【契約書未締結】オファー受諾済みだが契約未締結",
                "未入金":       "【未入金】契約書締結済みだが前払い未入金",
            }
            for key, label in category_labels.items():
                items = followups.get(key, [])
                if not items:
                    continue
                text.insert("end", f"{label}\n")
                for book, publisher, lang in items:
                    info = "／".join(filter(None, [publisher, lang]))
                    if info:
                        text.insert("end", f"  - {book}（{info}）\n")
                    else:
                        text.insert("end", f"  - {book}\n")
                text.insert("end", "\n")

        text.tag_config("section", font=section_font, foreground=p["title"])
        text.tag_config("section_warn", font=section_font, foreground="#e5a00d" if p["bg"].startswith("#1") or p["bg"].startswith("#0") or p["bg"] == "#2b2b2b" else "#b45309")
        text.configure(state="disabled")

        tk.Button(root, text="OK", command=root.destroy,
                  font=body_font, bg=p["btn_bg"], fg=p["btn_fg"],
                  relief="flat", padx=20, pady=5).pack(pady=10)

    build_ui(COLOR_PALETTES[palette_index])
    root.mainloop()


# ============================================================
# 📧 メール送信機能（1日1通制御つき）
# ============================================================

EMAIL_RECIPIENT = "matsmoto.norihito@gmail.com"


def is_already_sent_today():
    """今日すでにメール送信済みかチェック"""
    if not os.path.exists(LAST_SENT_FILE):
        return False
    try:
        with open(LAST_SENT_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        last_date = data.get("last_sent_date", "")
        today = datetime.date.today().isoformat()
        return last_date == today
    except Exception:
        return False


def mark_sent_today():
    """今日送信したことを記録"""
    today = datetime.date.today().isoformat()
    data = {
        "last_sent_date": today,
        "sent_from_host": HOSTNAME,
        "sent_at": datetime.datetime.now().isoformat(timespec="seconds"),
    }
    try:
        with open(LAST_SENT_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"送信記録エラー: {e}", file=sys.stderr)


def render_html_email(events, tasks, followups, palette):
    """HTMLメール本文を生成"""
    today = datetime.date.today()
    weekday_jp = ["月", "火", "水", "木", "金", "土", "日"][today.weekday()]
    today_str = f"{today.strftime('%Y年%m月%d日')}（{weekday_jp}）"

    bg = palette["bg"]
    fg = palette["fg"]
    title_color = palette["title"]
    accent = palette["btn_bg"]

    html = f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<style>
body {{ font-family: 'Hiragino Sans', 'Yu Gothic', 'Meiryo', sans-serif; background-color: {bg}; color: {fg}; margin: 0; padding: 20px; }}
.container {{ max-width: 600px; margin: 0 auto; background-color: {bg}; padding: 24px; border-radius: 12px; }}
h1 {{ color: {title_color}; font-size: 22px; border-bottom: 3px solid {accent}; padding-bottom: 8px; margin-top: 0; }}
h2 {{ color: {title_color}; font-size: 17px; margin-top: 28px; padding-left: 10px; border-left: 4px solid {accent}; }}
h3 {{ color: #e5a00d; font-size: 15px; margin-top: 20px; margin-bottom: 6px; }}
ul {{ list-style: none; padding-left: 0; margin: 8px 0; }}
li {{ padding: 6px 0 6px 20px; position: relative; line-height: 1.6; }}
li::before {{ content: "▸"; position: absolute; left: 0; color: {accent}; font-weight: bold; }}
.empty {{ color: #888; font-style: italic; padding-left: 20px; }}
.warning-section {{ background-color: rgba(229, 160, 13, 0.1); border-left: 4px solid #e5a00d; padding: 14px 18px; margin-top: 28px; border-radius: 6px; }}
.meta {{ color: #888; font-size: 12px; margin-top: 32px; padding-top: 14px; border-top: 1px solid rgba(128,128,128,0.3); }}
.task-list-name {{ color: {title_color}; font-weight: bold; margin-top: 14px; }}
.due {{ color: #e5a00d; font-size: 13px; }}
</style>
</head>
<body>
<div class="container">
<h1>📅 {today_str}</h1>
"""

    # カレンダー
    html += '<h2>🗓️ 今日の予定</h2>'
    if events:
        html += '<ul>'
        for ev in events:
            summary = ev.get("summary", "(無題)")
            start = ev.get("start", {})
            if "dateTime" in start:
                dt = datetime.datetime.fromisoformat(start["dateTime"].replace("Z", "+00:00"))
                time_str = dt.strftime("%H:%M")
                html += f'<li><strong>{time_str}</strong> &nbsp; {summary}</li>'
            else:
                html += f'<li><strong>終日</strong> &nbsp; {summary}</li>'
        html += '</ul>'
    else:
        html += '<div class="empty">予定はありません</div>'

    # タスク
    html += '<h2>✅ ToDoリスト</h2>'
    if tasks:
        # リスト名でグループ化
        from collections import defaultdict
        grouped = defaultdict(list)
        for task in tasks:
            grouped[task.get("_list_name", "未分類")].append(task)
        for list_name, items in grouped.items():
            html += f'<div class="task-list-name">{list_name}</div>'
            html += '<ul>'
            for task in items:
                title = task.get("title", "(無題)")
                due_str = ""
                if task.get("due"):
                    try:
                        d = datetime.datetime.fromisoformat(task["due"].replace("Z", "+00:00"))
                        due_str = f' <span class="due">(期限: {d.strftime("%m/%d")})</span>'
                    except Exception:
                        pass
                html += f'<li>{title}{due_str}</li>'
            html += '</ul>'
    else:
        html += '<div class="empty">タスクはありません</div>'

    # Airtableフォローアップ
    has_followups = any(items for items in followups.values()) if followups else False
    if has_followups:
        html += '<div class="warning-section">'
        html += '<h2 style="margin-top:0; border:none; padding:0;">⚠ Ascom 要フォローアップ</h2>'
        category_labels = {
            "社内判断待ち": "【社内判断待ち】オファーが出たまま社内判断保留中",
            "契約書未締結": "【契約書未締結】オファー受諾済みだが契約未締結",
            "未入金":       "【未入金】契約書締結済みだが前払い未入金",
        }
        for key, label in category_labels.items():
            items = followups.get(key, [])
            if not items:
                continue
            html += f'<h3>{label}</h3><ul>'
            for book, publisher, lang in items:
                info = " / ".join(filter(None, [publisher, lang]))
                if info:
                    html += f'<li>{book} <span style="color:#888;font-size:13px;">（{info}）</span></li>'
                else:
                    html += f'<li>{book}</li>'
            html += '</ul>'
        html += '</div>'

    html += f'<div class="meta">📍 送信元: {HOSTNAME} | 🕐 {datetime.datetime.now().strftime("%H:%M:%S")} 自動送信</div>'
    html += '</div></body></html>'
    return html


def send_email(creds, events, tasks, followups, palette):
    """Gmail APIで自分宛にHTMLメールを送信"""
    today = datetime.date.today()
    weekday_jp = ["月", "火", "水", "木", "金", "土", "日"][today.weekday()]
    subject = f"【Daily Notice】{today.isoformat()} ({weekday_jp}) 今日のタスク・予定"

    html_body = render_html_email(events, tasks, followups, palette)

    message = MIMEText(html_body, "html", "utf-8")
    message["to"] = EMAIL_RECIPIENT
    message["from"] = EMAIL_RECIPIENT
    message["subject"] = subject

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")

    try:
        service = build("gmail", "v1", credentials=creds)
        service.users().messages().send(userId="me", body={"raw": raw}).execute()
        print(f"✉️ メール送信完了: {EMAIL_RECIPIENT}")
        mark_sent_today()
        return True
    except Exception as e:
        print(f"メール送信エラー: {e}", file=sys.stderr)
        return False


def main():
    force_choose = "--choose-color" in sys.argv
    demo_mode    = "--demo" in sys.argv

    config = load_config()
    palette_index = config.get("palette_index", 0)

    # --choose-color のときは仮ウィンドウでパレット選択
    if force_choose:
        tmp = tk.Tk()
        tmp.withdraw()
        palette_index = choose_color_palette(tmp, current_index=palette_index)
        tmp.destroy()
        config["palette_index"] = palette_index
        save_config(config)

    if demo_mode:
        events = get_demo_events()
        tasks  = get_demo_tasks()
    else:
        creds  = get_credentials()
        events = get_calendar_events(creds)
        tasks  = get_tasks(creds)

    # Airtableフォローアップ取得
    followups = get_airtable_followups(config)

    show_popup(events, tasks, palette_index, followups)

    # ============================================================
    # 📧 1日1通メール送信（demo時はスキップ、本日送信済みならスキップ）
    # ============================================================
    if not demo_mode and not is_already_sent_today():
        try:
            send_email(creds, events, tasks, followups, COLOR_PALETTES[palette_index])
        except Exception as e:
            print(f"メール送信処理エラー: {e}", file=sys.stderr)


if __name__ == "__main__":
    main()
