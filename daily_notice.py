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
#    名前・背景色・文字色・タイトル色・ボタン背景・ボタン文字を設定
# ============================================================
COLOR_PALETTES = [
    {
        "name": "🌙 ダーク（デフォルト）",
        "bg": "#1e1e2e",
        "fg": "#cdd6f4",
        "title": "#89b4fa",
        "btn_bg": "#89b4fa",
        "btn_fg": "#1e1e2e",
    },
    {
        "name": "☁️ クリーンホワイト",
        "bg": "#ffffff",
        "fg": "#1a1a1a",
        "title": "#2d6a4f",
        "btn_bg": "#2d6a4f",
        "btn_fg": "#ffffff",
    },
    {
        "name": "🌿 ナチュラルグリーン",
        "bg": "#e8f5e9",
        "fg": "#1b5e20",
        "title": "#388e3c",
        "btn_bg": "#388e3c",
        "btn_fg": "#ffffff",
    },
    {
        "name": "🌅 ウォームベージュ",
        "bg": "#fdf6ec",
        "fg": "#3e2723",
        "title": "#bf360c",
        "btn_bg": "#bf360c",
        "btn_fg": "#ffffff",
    },
    {
        "name": "🌌 ミッドナイトブルー",
        "bg": "#0d1b2a",
        "fg": "#e0e0e0",
        "title": "#90caf9",
        "btn_bg": "#1565c0",
        "btn_fg": "#ffffff",
    },
    {
        "name": "🌸 ソフトピンク",
        "bg": "#fce4ec",
        "fg": "#4a0030",
        "title": "#c2185b",
        "btn_bg": "#c2185b",
        "btn_fg": "#ffffff",
    },
    {
        "name": "🌞 サンシャインイエロー",
        "bg": "#fffde7",
        "fg": "#3e2700",
        "title": "#f57f17",
        "btn_bg": "#f57f17",
        "btn_fg": "#ffffff",
    },
    {
        "name": "🪨 モノクロームグレー",
        "bg": "#2b2b2b",
        "fg": "#e0e0e0",
        "title": "#bdbdbd",
        "btn_bg": "#616161",
        "btn_fg": "#ffffff",
    },
]
# ============================================================


def load_config():
    """保存済みの設定を読み込む。なければデフォルト（パレット0番）を返す。"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"palette_index": 0}


def save_config(config):
    """設定をファイルに保存する。"""
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


def choose_color_palette(current_index=0):
    """
    カラーパレット選択ダイアログを表示する。
    選択されたパレットのインデックスを返す。キャンセル時は current_index を返す。
    """
    selected_index = [current_index]

    dialog = tk.Tk()
    dialog.title("🎨 ポップアップの色テーマを選んでください")
    dialog.geometry("400x480")
    dialog.resizable(False, False)
    dialog.configure(bg="#f5f5f5")

    tk.Label(
        dialog,
        text="色テーマを選んでください",
        font=("Meiryo UI", 13, "bold"),
        bg="#f5f5f5",
        fg="#1a1a1a",
        pady=14,
    ).pack()

    frame = tk.Frame(dialog, bg="#f5f5f5")
    frame.pack(fill="both", expand=True, padx=20)

    var = tk.IntVar(value=current_index)

    def update_preview(idx):
        p = COLOR_PALETTES[idx]
        preview_label.configure(
            bg=p["bg"],
            fg=p["fg"],
            text=f'背景: {p["bg"]}　文字: {p["fg"]}',
        )
        preview_title.configure(bg=p["bg"], fg=p["title"])

    for i, palette in enumerate(COLOR_PALETTES):
        rb = tk.Radiobutton(
            frame,
            text=palette["name"],
            variable=var,
            value=i,
            font=("Meiryo UI", 11),
            bg="#f5f5f5",
            activebackground="#e0e0e0",
            command=lambda idx=i: update_preview(idx),
            anchor="w",
        )
        rb.pack(fill="x", pady=2)

    # プレビューエリア
    preview_frame = tk.Frame(dialog, bg="#f5f5f5")
    preview_frame.pack(fill="x", padx=20, pady=(6, 2))
    p0 = COLOR_PALETTES[current_index]
    preview_title = tk.Label(
        preview_frame,
        text="2026年04月10日(Fri)  ← プレビュー",
        font=("Meiryo UI", 10, "bold"),
        bg=p0["bg"],
        fg=p0["title"],
        padx=10,
        pady=4,
        anchor="w",
    )
    preview_title.pack(fill="x")
    preview_label = tk.Label(
        preview_frame,
        text=f'背景: {p0["bg"]}　文字: {p0["fg"]}',
        font=("Meiryo UI", 10),
        bg=p0["bg"],
        fg=p0["fg"],
        padx=10,
        pady=4,
        anchor="w",
    )
    preview_label.pack(fill="x")

    def on_ok():
        selected_index[0] = var.get()
        dialog.destroy()

    def on_cancel():
        dialog.destroy()

    btn_frame = tk.Frame(dialog, bg="#f5f5f5")
    btn_frame.pack(pady=12)
    tk.Button(
        btn_frame, text="この色に決定", command=on_ok,
        font=("Meiryo UI", 11, "bold"),
        bg="#2d6a4f", fg="#ffffff", relief="flat", padx=16, pady=6,
    ).pack(side="left", padx=8)
    tk.Button(
        btn_frame, text="キャンセル", command=on_cancel,
        font=("Meiryo UI", 11),
        bg="#9e9e9e", fg="#ffffff", relief="flat", padx=16, pady=6,
    ).pack(side="left", padx=8)

    dialog.mainloop()
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
    end_of_day = now.replace(hour=23, minute=59, second=59, microsecond=0).isoformat() + "+09:00"

    events_result = service.events().list(
        calendarId="primary",
        timeMin=start_of_day,
        timeMax=end_of_day,
        singleEvents=True,
        orderBy="startTime",
    ).execute()
    return events_result.get("items", [])


def get_tasks(creds):
    service = build("tasks", "v1", credentials=creds)
    tasklists = service.tasklists().list().execute().get("items", [])
    all_tasks = []
    for tasklist in tasklists:
        tasks = (
            service.tasks()
            .list(tasklist=tasklist["id"], showCompleted=False, showHidden=False)
            .execute()
            .get("items", [])
        )
        for task in tasks:
            task["_list_name"] = tasklist["title"]
        all_tasks.extend(tasks)
    return all_tasks


def format_time(event):
    start = event["start"].get("dateTime", event["start"].get("date", ""))
    if "T" in start:
        t = datetime.datetime.fromisoformat(start)
        return t.strftime("%H:%M")
    return "終日"


def show_popup(events, tasks, palette):
    today = datetime.date.today().strftime("%Y年%m月%d日(%a)")

    root = tk.Tk()
    root.title(f"Today's Notice - {today}")
    root.geometry("520x600")
    root.attributes("-topmost", True)
    root.configure(bg=palette["bg"])

    title_font = ("Meiryo UI", 14, "bold")
    section_font = ("Meiryo UI", 12, "bold")
    body_font = ("Meiryo UI", 11)

    # タイトル行（右端に「色を変える」ボタン）
    header_frame = tk.Frame(root, bg=palette["bg"])
    header_frame.pack(fill="x")

    tk.Label(
        header_frame, text=f"{today}", font=title_font,
        bg=palette["bg"], fg=palette["fg"], pady=10, anchor="w", padx=15,
    ).pack(side="left", fill="x", expand=True)

    def on_change_color():
        config = load_config()
        new_index = choose_color_palette(current_index=config.get("palette_index", 0))
        config["palette_index"] = new_index
        save_config(config)
        root.destroy()
        # 新しい配色で再表示
        show_popup(events, tasks, COLOR_PALETTES[new_index])

    tk.Button(
        header_frame, text="🎨 色変更", command=on_change_color,
        font=("Meiryo UI", 9), bg=palette["btn_bg"], fg=palette["btn_fg"],
        relief="flat", padx=8, pady=4,
    ).pack(side="right", padx=10, pady=8)

    # スクロール可能テキスト
    text = scrolledtext.ScrolledText(
        root, font=body_font, bg=palette["bg"], fg=palette["fg"],
        insertbackground=palette["fg"], wrap="word", relief="flat",
        padx=15, pady=10,
    )
    text.pack(fill="both", expand=True)

    # カレンダー予定
    text.insert("end", "--- Google Calendar ---\n\n", "section")
    if events:
        for event in events:
            time_str = format_time(event)
            summary = event.get("summary", "(無題)")
            text.insert("end", f"  {time_str}  {summary}\n")
    else:
        text.insert("end", "  予定はありません\n")

    text.insert("end", "\n")

    # タスク
    text.insert("end", "--- ToDo List ---\n\n", "section")
    if tasks:
        current_list = ""
        for task in tasks:
            list_name = task.get("_list_name", "")
            if list_name != current_list:
                text.insert("end", f"  [{list_name}]\n")
                current_list = list_name
            title = task.get("title", "(無題)")
            due = task.get("due", "")
            due_str = ""
            if due:
                d = datetime.datetime.fromisoformat(due.replace("Z", "+00:00"))
                due_str = f" (期限: {d.strftime('%m/%d')})"
            text.insert("end", f"    - {title}{due_str}\n")
    else:
        text.insert("end", "  タスクはありません\n")

    text.tag_config("section", font=section_font, foreground=palette["title"])
    text.configure(state="disabled")

    # 閉じるボタン
    tk.Button(
        root, text="OK", command=root.destroy,
        font=body_font, bg=palette["btn_bg"], fg=palette["btn_fg"],
        relief="flat", padx=20, pady=5,
    ).pack(pady=10)

    root.mainloop()


def main():
    # --choose-color 引数があれば色選択ダイアログを先に開く
    force_choose = "--choose-color" in sys.argv

    config = load_config()
    palette_index = config.get("palette_index", 0)

    if force_choose:
        palette_index = choose_color_palette(current_index=palette_index)
        config["palette_index"] = palette_index
        save_config(config)

    palette = COLOR_PALETTES[palette_index]

    creds = get_credentials()
    events = get_calendar_events(creds)
    tasks = get_tasks(creds)
    show_popup(events, tasks, palette)


if __name__ == "__main__":
    main()
