"""
Teams Chat Exporter v4.0
Получение токена: CDP (Chrome DevTools Protocol) — подключение к запущенному Edge/Chrome
Требования: pip install requests websocket-client
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import requests
import json
import csv
import time
import os
import re
import urllib3
import webbrowser
import subprocess
import socket

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

DEFAULT_OUTPUT = r"C:\Users\AEvstratov\Downloads"
GRAPH_BASE     = "https://graph.microsoft.com/v1.0"
TEAMS_COLOR    = "#6264a7"
CDP_PORT       = 9222   # стандартный порт отладки

# ─── CDP HELPERS ──────────────────────────────────────────────────────────────

def find_edge_exe():
    candidates = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return None

def find_chrome_exe():
    candidates = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expanduser(r"~\AppData\Local\Google\Chrome\Application\chrome.exe"),
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return None

def is_cdp_running():
    try:
        s = socket.create_connection(("127.0.0.1", CDP_PORT), timeout=1)
        s.close()
        return True
    except:
        return False

def launch_browser_with_cdp(url):
    """Запускает Edge или Chrome с отладочным портом"""
    browser = find_edge_exe() or find_chrome_exe()
    if not browser:
        return None, "Edge и Chrome не найдены"

    user_data = os.path.join(os.environ.get("LOCALAPPDATA",""), "TeamsExporterCDP")
    cmd = [
        browser,
        f"--remote-debugging-port={CDP_PORT}",
        f"--user-data-dir={user_data}",
        "--no-first-run",
        "--no-default-browser-check",
        url,
    ]
    proc = subprocess.Popen(cmd)
    # Ждём запуска CDP
    for _ in range(20):
        if is_cdp_running():
            return proc, None
        time.sleep(0.5)
    return proc, "CDP не запустился — браузер не ответил"

def cdp_get_tabs():
    try:
        r = requests.get(f"http://127.0.0.1:{CDP_PORT}/json", timeout=3)
        return r.json()
    except:
        return []

def cdp_eval(ws_url, expression, timeout=10):
    """Выполняет JS в вкладке через WebSocket CDP"""
    try:
        import websocket
    except ImportError:
        raise Exception("Установи: pip install websocket-client")

    result = {"value": None, "error": None}
    done = threading.Event()

    ws = websocket.WebSocket()
    ws.connect(ws_url, timeout=timeout)

    msg_id = 1
    ws.send(json.dumps({
        "id": msg_id,
        "method": "Runtime.evaluate",
        "params": {
            "expression": expression,
            "returnByValue": True,
            "awaitPromise": False,
        }
    }))

    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            ws.settimeout(2)
            raw = ws.recv()
            data = json.loads(raw)
            if data.get("id") == msg_id:
                r = data.get("result", {})
                if "exceptionDetails" in r:
                    result["error"] = str(r["exceptionDetails"])
                else:
                    result["value"] = r.get("result", {}).get("value")
                break
        except:
            break
    ws.close()
    return result["value"]

JS_GET_TOKEN = """
(function() {
    var k = Object.keys(localStorage).find(function(k) {
        return k.includes('accesstoken') && k.includes('graph.microsoft.com');
    });
    if (!k) return null;
    try {
        var v = JSON.parse(localStorage.getItem(k));
        if (v && v.secret && v.secret.length > 100) {
            var exp = parseInt(v.expiresOn || v.extendedExpiresOn || 0);
            if (exp > Date.now()/1000) return v.secret;
        }
    } catch(e) {}
    return null;
})()
"""

GE_URL = "https://developer.microsoft.com/en-us/graph/graph-explorer"

def find_ge_tab(tabs):
    """Ищет вкладку Graph Explorer среди открытых"""
    for tab in tabs:
        url = tab.get("url", "")
        if "graph-explorer" in url or "developer.microsoft.com" in url:
            return tab
    return None

# ─── GRAPH API ────────────────────────────────────────────────────────────────

def api_get(url, token):
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"}, verify=False)
    if resp.status_code == 401:
        raise Exception("Токен истёк или недействителен")
    if resp.status_code == 403:
        raise Exception(f"Нет доступа: {resp.text[:200]}")
    resp.raise_for_status()
    return resp.json()

def fetch_chats(token):
    data = api_get(f"{GRAPH_BASE}/me/chats?$expand=members&$top=50", token)
    chats = []
    for c in data.get("value", []):
        topic = c.get("topic") or ""
        members = c.get("members", [])
        if not topic:
            names = [m.get("displayName", "?") for m in members[:3]]
            topic = ", ".join(names) or "(без названия)"
        chats.append({"id": c["id"], "topic": topic, "type": c.get("chatType", "")})
    return chats

def fetch_messages(chat_id, token, progress_cb=None):
    url = f"{GRAPH_BASE}/chats/{chat_id}/messages?$top=50"
    messages = []
    page = 0
    while url:
        page += 1
        if progress_cb:
            progress_cb(page, len(messages))
        data = api_get(url, token)
        messages.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
        time.sleep(0.3)
    return messages

def extract_text(body):
    if not body: return ""
    c = body.get("content", "") or ""
    if body.get("contentType") == "html":
        c = re.sub(r'<br\s*/?>', '\n', c)
        c = re.sub(r'</p>', '\n', c)
        c = re.sub(r'<[^>]+>', '', c)
        c = c.replace('&nbsp;',' ').replace('&lt;','<').replace('&gt;','>').replace('&amp;','&')
        c = re.sub(r'\n{3,}', '\n\n', c).strip()
    return c

def extract_message_refs(msg):
    refs = []
    for att in (msg.get("attachments") or []):
        if att.get("contentType") == "messageReference":
            try:
                rc = json.loads(att.get("content") or "{}")
                sender = rc.get("messageSender", {})
                name = (sender.get("user") or sender.get("application") or {}).get("displayName","")
                preview = rc.get("messagePreview") or rc.get("body",{}).get("content","")
                preview = re.sub(r'<[^>]+>', '', preview or "")
                preview = preview.replace('&nbsp;',' ').replace('&lt;','<').replace('&gt;','>').strip()
                if preview:
                    refs.append({"sender": name, "preview": preview[:200]})
            except: pass
    return refs

def parse_msg(msg):
    sender = msg.get("from") or {}
    user   = sender.get("user") or sender.get("application") or {}
    name   = user.get("displayName") or "System"
    raw    = msg.get("createdDateTime", "")
    try:
        from datetime import datetime
        dt    = datetime.fromisoformat(raw.replace("Z","+00:00"))
        ts    = dt.strftime("%Y-%m-%d %H:%M:%S")
        date  = dt.strftime("%Y-%m-%d")
        ttime = dt.strftime("%H:%M:%S")
    except:
        ts = raw; date = ""; ttime = ""
    return {
        "id":          msg.get("id",""),
        "created":     ts, "date": date, "time": ttime,
        "sender":      name,
        "text":        extract_text(msg.get("body")),
        "attachments": ", ".join(
            str(a.get("name") or "file")
            for a in (msg.get("attachments") or [])
            if a.get("contentType") != "messageReference" and a.get("name")
        ),
        "msg_refs":    extract_message_refs(msg),
        "deleted":     msg.get("deletedDateTime") is not None,
        "type":        msg.get("messageType",""),
        "importance":  msg.get("importance","normal"),
        "reply_to_id": msg.get("replyToId") or "",
    }

def save_exports(parsed, raw, topic, output_dir):
    from datetime import datetime
    safe  = re.sub(r'[\\/:*?"<>|]','_', topic)[:60]
    ts    = datetime.now().strftime("%Y%m%d_%H%M%S")
    base  = os.path.join(output_dir, f"teams_{safe}_{ts}")

    csv_path = base + ".csv"
    fields   = [k for k in parsed[0].keys() if k != "msg_refs"]
    with open(csv_path,"w",encoding="utf-8-sig",newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        w.writeheader(); w.writerows(parsed)

    json_path = base + ".json"
    with open(json_path,"w",encoding="utf-8") as f:
        json.dump(raw, f, ensure_ascii=False, indent=2)

    html_path  = base + ".html"
    rows       = []
    prev_date  = None
    msg_index  = {m["id"]: m for m in parsed if "id" in m}

    for m in reversed(parsed):
        if m["deleted"] or m["type"] not in ("message",""):
            continue
        if m["date"] != prev_date:
            rows.append(f'<div class="ds">{m["date"]}</div>')
            prev_date = m["date"]

        txt = (m["text"]
               .replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
               .replace("\n","<br>"))
        att = f'<div class="at">📎 {m["attachments"]}</div>' if m["attachments"] else ""
        imp = ' style="border-left:3px solid #e8a000"' if m["importance"]=="high" else ""

        quote_html = ""
        rid = m.get("reply_to_id","")
        if rid and rid in msg_index:
            ref     = msg_index[rid]
            ref_txt = ref["text"][:150].replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace("\n"," ")
            if len(ref["text"]) > 150: ref_txt += "…"
            quote_html = f'<div class="qt"><span class="qs">{ref["sender"]}</span>: {ref_txt}</div>'
        elif m.get("msg_refs"):
            parts = []
            for ref in m["msg_refs"]:
                pv = ref["preview"][:150].replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
                if len(ref["preview"])>150: pv += "…"
                parts.append(f'<span class="qs">{ref["sender"] or "?"}</span>: {pv}')
            quote_html = f'<div class="qt">{" / ".join(parts)}</div>'

        rows.append(
            f'<div class="m"{imp}>'
            f'<div class="h"><span class="s">{m["sender"]}</span><span class="t">{m["time"]}</span></div>'
            f'{quote_html}<div class="b">{txt}</div>{att}</div>'
        )

    visible = len([m for m in parsed if not m["deleted"]])
    from datetime import datetime
    html = f"""<!DOCTYPE html><html lang="ru"><head><meta charset="UTF-8">
<title>{topic}</title><style>
*{{box-sizing:border-box}}
body{{font-family:'Segoe UI',sans-serif;background:#f3f2f1;margin:0;padding:0}}
.header{{background:{TEAMS_COLOR};color:#fff;padding:16px 24px;position:sticky;top:0;z-index:10;box-shadow:0 2px 8px rgba(0,0,0,.2)}}
.header h1{{margin:0;font-size:16px;font-weight:600}}
.header .meta{{font-size:12px;opacity:.8;margin-top:2px}}
.container{{max-width:860px;margin:0 auto;padding:16px}}
.m{{background:#fff;border-radius:8px;padding:10px 14px;margin:4px 0;box-shadow:0 1px 2px rgba(0,0,0,.06);transition:box-shadow .2s}}
.m:hover{{box-shadow:0 2px 8px rgba(0,0,0,.12)}}
.h{{display:flex;gap:8px;align-items:baseline;margin-bottom:4px}}
.s{{font-weight:600;color:{TEAMS_COLOR};font-size:13px}}
.t{{color:#aaa;font-size:11px}}
.b{{font-size:14px;color:#252423;line-height:1.55;word-break:break-word}}
.at{{font-size:12px;color:#888;margin-top:4px}}
.ds{{text-align:center;color:#999;font-size:11px;margin:16px 0 6px;display:flex;align-items:center;gap:8px}}
.ds::before,.ds::after{{content:'';flex:1;height:1px;background:#e0e0e0}}
.qt{{background:#f0f0f8;border-left:3px solid {TEAMS_COLOR};padding:4px 8px;border-radius:4px;
    font-size:12px;color:#555;margin-bottom:6px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
.qs{{font-weight:600;color:{TEAMS_COLOR}}}
</style></head><body>
<div class="header"><h1>💬 {topic}</h1>
<div class="meta">Экспорт: {datetime.now().strftime('%d.%m.%Y %H:%M')} · {visible} сообщений</div></div>
<div class="container">{''.join(rows)}</div></body></html>"""

    with open(html_path,"w",encoding="utf-8") as f:
        f.write(html)
    return csv_path, json_path, html_path

# ─── GUI ──────────────────────────────────────────────────────────────────────

class TeamsExporterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Teams Chat Exporter v4.0")
        self.root.geometry("720x680")
        self.root.resizable(True, True)
        self.root.configure(bg="#f3f2f1")
        self.token        = tk.StringVar()
        self.chats        = []
        self.output_dir   = tk.StringVar(value=DEFAULT_OUTPUT)
        self.selected_chat_id    = None
        self.selected_chat_topic = None
        self._browser_proc = None
        self._polling      = False
        self._build_ui()

    def _build_ui(self):
        hdr = tk.Frame(self.root, bg=TEAMS_COLOR, height=56)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text="  💬  Teams Chat Exporter",
                 bg=TEAMS_COLOR, fg="white",
                 font=("Segoe UI",14,"bold")).pack(side="left", padx=8, pady=12)

        main = tk.Frame(self.root, bg="#f3f2f1")
        main.pack(fill="both", expand=True, padx=16, pady=12)

        self._section(main, "Шаг 1 — Токен (Graph Explorer)")

        # Кнопка автозахвата
        self.btn_cdp = tk.Button(main,
            text="🔑  Открыть Graph Explorer и захватить токен автоматически",
            command=self._start_cdp_capture,
            bg=TEAMS_COLOR, fg="white", relief="flat",
            font=("Segoe UI",9,"bold"), cursor="hand2",
            activebackground="#5052a0", activeforeground="white")
        self.btn_cdp.pack(fill="x", ipady=7, pady=(0,6))

        # Инструкция
        info = tk.Frame(main, bg="#e8eaf6")
        info.pack(fill="x", pady=(0,6))
        tk.Label(info,
            text="  ℹ️  Нажми кнопку → войди в Graph Explorer → токен захватится автоматически",
            bg="#e8eaf6", fg="#444", font=("Segoe UI",8), anchor="w").pack(fill="x", pady=4)

        # Ручной ввод токена
        token_row = tk.Frame(main, bg="#f3f2f1")
        token_row.pack(fill="x", pady=(0,8))
        self.token_entry = tk.Entry(token_row, textvariable=self.token,
            show="•", font=("Segoe UI",9), relief="solid", bd=1)
        self.token_entry.pack(side="left", fill="x", expand=True, ipady=5, padx=(0,5))
        tk.Button(token_row, text="👁", command=self._toggle_token,
            bg="white", relief="solid", bd=1, font=("Segoe UI",9),
            cursor="hand2", width=3).pack(side="left", padx=(0,3))
        tk.Button(token_row, text="📋 Вставить", command=self._paste_token,
            bg="white", relief="solid", bd=1, font=("Segoe UI",9),
            cursor="hand2").pack(side="left")

        # Кнопка загрузки чатов
        tk.Button(main, text="🔍  Загрузить список чатов",
            command=self._load_chats,
            bg=TEAMS_COLOR, fg="white", relief="flat",
            font=("Segoe UI",10,"bold"), cursor="hand2",
            activebackground="#5052a0", activeforeground="white"
        ).pack(fill="x", ipady=8, pady=(0,12))

        self._section(main, "Шаг 2 — Выберите чат")
        list_frame = tk.Frame(main, bg="#f3f2f1")
        list_frame.pack(fill="both", expand=True, pady=(0,10))
        sb = ttk.Scrollbar(list_frame); sb.pack(side="right", fill="y")
        self.chat_listbox = tk.Listbox(list_frame,
            yscrollcommand=sb.set, font=("Segoe UI",10),
            selectbackground=TEAMS_COLOR, selectforeground="white",
            relief="solid", bd=1, activestyle="none", height=8)
        self.chat_listbox.pack(side="left", fill="both", expand=True)
        sb.config(command=self.chat_listbox.yview)
        self.chat_listbox.bind("<<ListboxSelect>>", self._on_chat_select)

        self._section(main, "Шаг 3 — Папка для сохранения")
        dir_row = tk.Frame(main, bg="#f3f2f1")
        dir_row.pack(fill="x", pady=(0,10))
        tk.Entry(dir_row, textvariable=self.output_dir,
            font=("Segoe UI",9), relief="solid", bd=1
        ).pack(side="left", fill="x", expand=True, ipady=5, padx=(0,5))
        tk.Button(dir_row, text="📁", command=self._browse_dir,
            bg="white", relief="solid", bd=1, font=("Segoe UI",10),
            cursor="hand2").pack(side="left")

        self.btn_export = tk.Button(main, text="⬇️  Экспортировать чат",
            command=self._start_export,
            bg="#107c10", fg="white", relief="flat",
            font=("Segoe UI",11,"bold"), cursor="hand2", state="disabled",
            activebackground="#0a5c0a", activeforeground="white")
        self.btn_export.pack(fill="x", ipady=10)

        status_bar = tk.Frame(self.root, bg="#e0e0e0", height=28)
        status_bar.pack(fill="x", side="bottom"); status_bar.pack_propagate(False)
        self.progress = ttk.Progressbar(status_bar, mode="indeterminate", length=120)
        self.progress.pack(side="right", padx=8, pady=4)
        self.status_var = tk.StringVar(value="Готов к работе")
        tk.Label(status_bar, textvariable=self.status_var,
            bg="#e0e0e0", fg="#444", font=("Segoe UI",9), anchor="w"
        ).pack(side="left", fill="x", padx=8)

    def _section(self, parent, text):
        f = tk.Frame(parent, bg="#f3f2f1"); f.pack(fill="x", pady=(4,4))
        tk.Label(f, text=text, bg="#f3f2f1", fg=TEAMS_COLOR,
                 font=("Segoe UI",9,"bold")).pack(anchor="w")
        tk.Frame(f, bg=TEAMS_COLOR, height=1).pack(fill="x", pady=(2,6))

    # ── CDP захват токена ──────────────────────────────────────────────────────

    def _start_cdp_capture(self):
        if self._polling:
            self._set_status("Уже идёт захват токена...")
            return
        self._polling = True
        self.btn_cdp.config(state="disabled", text="⏳  Ожидаю вход в Graph Explorer...")
        self.progress.start(10)
        threading.Thread(target=self._cdp_thread, daemon=True).start()

    def _cdp_thread(self):
        try:
            # 1. Запускаем браузер с CDP если ещё не запущен
            if not is_cdp_running():
                self.root.after(0, self._set_status, "Запускаю браузер с отладочным портом...")
                proc, err = launch_browser_with_cdp(GE_URL)
                if err:
                    self.root.after(0, self._cdp_error, err)
                    return
                self._browser_proc = proc
                time.sleep(2)
            else:
                # CDP уже запущен — открываем новую вкладку Graph Explorer
                self.root.after(0, self._set_status, "CDP уже запущен, открываю Graph Explorer...")
                try:
                    requests.get(f"http://127.0.0.1:{CDP_PORT}/json/new?"
                                 + GE_URL, timeout=3)
                except:
                    webbrowser.open(GE_URL)

            self.root.after(0, self._set_status,
                "Войдите в Graph Explorer в открывшемся браузере...")

            # 2. Polling — ждём появления токена
            for attempt in range(150):  # ~5 минут
                time.sleep(2)
                tabs = cdp_get_tabs()
                ge_tab = find_ge_tab(tabs)
                if not ge_tab:
                    continue
                ws_url = ge_tab.get("webSocketDebuggerUrl")
                if not ws_url:
                    continue
                try:
                    token = cdp_eval(ws_url, JS_GET_TOKEN, timeout=5)
                    if token and len(token) > 100:
                        self.root.after(0, self._on_token_captured, token)
                        return
                except Exception:
                    pass
                if attempt % 5 == 0:
                    self.root.after(0, self._set_status,
                        f"Жду токен... ({attempt*2}с) — войдите в Graph Explorer")

            self.root.after(0, self._cdp_error, "Время ожидания истекло (5 мин)")

        except Exception as e:
            self.root.after(0, self._cdp_error, str(e))

    def _on_token_captured(self, token):
        self._polling = False
        self.progress.stop()
        self.token.set(token)
        self.btn_cdp.config(state="normal",
            text="✅  Токен получен! Нажми снова для обновления")
        self._set_status("✅ Токен захвачен автоматически")
        messagebox.showinfo("Готово",
            "Токен успешно захвачен из Graph Explorer!\n\n"
            "Теперь нажми 'Загрузить список чатов'.")

    def _cdp_error(self, msg):
        self._polling = False
        self.progress.stop()
        self.btn_cdp.config(state="normal",
            text="🔑  Открыть Graph Explorer и захватить токен автоматически")
        self._set_status(f"❌ {msg[:80]}")
        messagebox.showerror("Ошибка CDP", msg + "\n\nВставь токен вручную через кнопку 📋")

    # ── Остальные методы ──────────────────────────────────────────────────────

    def _toggle_token(self):
        self.token_entry.config(
            show="" if self.token_entry.cget("show")=="•" else "•")

    def _paste_token(self):
        try:
            t = self.root.clipboard_get().strip()
            if t.startswith("Bearer "): t = t[7:]
            self.token.set(t)
            self._set_status("Токен вставлен из буфера")
        except:
            self._set_status("Буфер обмена пуст")

    def _browse_dir(self):
        d = filedialog.askdirectory(initialdir=self.output_dir.get())
        if d: self.output_dir.set(d)

    def _set_status(self, msg):
        self.status_var.set(msg)
        self.root.update_idletasks()

    def _load_chats(self):
        token = self.token.get().strip()
        if not token:
            messagebox.showwarning("Токен", "Сначала получи токен")
            return
        self._set_status("Загружаю список чатов...")
        self.progress.start(10)
        threading.Thread(target=self._load_chats_thread, args=(token,), daemon=True).start()

    def _load_chats_thread(self, token):
        try:
            chats = fetch_chats(token)
            self.chats = chats
            self.root.after(0, self._update_chat_list, chats)
        except Exception as e:
            self.root.after(0, self._show_error, str(e))

    def _update_chat_list(self, chats):
        self.progress.stop()
        self.chat_listbox.delete(0, tk.END)
        for c in chats:
            icon = "👥" if c["type"]=="group" else "💬"
            self.chat_listbox.insert(tk.END, f"  {icon}  {c['topic']}")
        self._set_status(f"Загружено {len(chats)} чатов — выберите нужный")

    def _on_chat_select(self, event):
        sel = self.chat_listbox.curselection()
        if sel:
            idx = sel[0]
            self.selected_chat_id    = self.chats[idx]["id"]
            self.selected_chat_topic = self.chats[idx]["topic"]
            self.btn_export.config(state="normal")
            self._set_status(f"Выбран: {self.selected_chat_topic}")

    def _start_export(self):
        token = self.token.get().strip()
        if not token:
            messagebox.showwarning("Токен","Токен не введён"); return
        if not self.selected_chat_id:
            messagebox.showwarning("Чат","Выберите чат"); return
        out = self.output_dir.get()
        if not os.path.isdir(out):
            messagebox.showwarning("Папка",f"Папка не существует:\n{out}"); return
        self.btn_export.config(state="disabled")
        self.progress.start(10)
        threading.Thread(target=self._export_thread,
            args=(token, self.selected_chat_id, self.selected_chat_topic, out),
            daemon=True).start()

    def _export_thread(self, token, chat_id, topic, output_dir):
        try:
            def pcb(page, count):
                self.root.after(0, self._set_status,
                    f"Страница {page}, загружено {count} сообщений...")
            raw    = fetch_messages(chat_id, token, pcb)
            self.root.after(0, self._set_status, "Обрабатываю и сохраняю...")
            parsed = [parse_msg(m) for m in raw]
            paths  = save_exports(parsed, raw, topic, output_dir)
            self.root.after(0, self._export_done, topic, len(parsed), *paths)
        except Exception as e:
            self.root.after(0, self._show_error, str(e))

    def _export_done(self, topic, count, html_path, csv_path, json_path):
        self.progress.stop()
        self.btn_export.config(state="normal")
        self._set_status(f"✅ Готово! {count} сообщений")
        if messagebox.askyesno("Готово!",
            f"✅ Чат «{topic}» экспортирован!\n\n"
            f"Сообщений: {count}\n"
            f"Папка: {os.path.dirname(html_path)}\n\n"
            f"Открыть HTML файл?"):
            os.startfile(html_path)

    def _show_error(self, msg):
        self.progress.stop()
        self.btn_export.config(state="normal")
        self._set_status(f"❌ {msg[:80]}")
        messagebox.showerror("Ошибка", msg)

# ─── MAIN ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    root = tk.Tk()
    try: root.tk.call('tk','scaling',1.25)
    except: pass
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("TScrollbar", background="#c0c0c0", troughcolor="#f3f2f1")
    style.configure("TProgressbar", background=TEAMS_COLOR)
    app = TeamsExporterApp(root)
    root.mainloop()
