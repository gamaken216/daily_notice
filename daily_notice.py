"""
PC起動時に今日のGoogleカレンダー予定とToDoリストをポップアップ表示するスクリプト
"""

import os
import sys
import json
import datetime
import tkinter as tk
from tkinter import scrolledtext
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = [
    "https://www.googleapis.com/auth/calendar.readonly",
    "https://www.googleapis.com/auth/tasks.readonly",
]

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_FILE = os.path.join(SCRIPT_DIR, "credentials.json")
TOKEN_FILE = os.path.join(SCRIPT_DIR, "token.json")
CONFIG_FILE = os.path.join(SCRIPT_DIR, "config.json")

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


def show_popup(events, tasks, palette_index):
    """メインポップアップ。palette_index を保持し、色変更時はウィジェットを再描画する。"""
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

        text.tag_config("section", font=section_font, foreground=p["title"])
        text.configure(state="disabled")

        tk.Button(root, text="OK", command=root.destroy,
                  font=body_font, bg=p["btn_bg"], fg=p["btn_fg"],
                  relief="flat", padx=20, pady=5).pack(pady=10)

    build_ui(COLOR_PALETTES[palette_index])
    root.mainloop()


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

    show_popup(events, tasks, palette_index)


if __name__ == "__main__":
    main()
