# -*- coding: utf-8 -*-
"""
桌面悬浮工作台 · 快捷复制 + 验证码助手 + 分组
现代卡片风格；分组管理；平滑滚动；邮箱/短信验证码
"""
import base64
import ctypes
import ctypes.wintypes
import email
import imaplib
import json
import os
import re
import shutil
import socket
import subprocess
import threading
import time
import tkinter as tk
import tkinter.ttk as ttk
import urllib.request
import zipfile
from contextlib import contextmanager
from email.header import decode_header
from email.utils import parsedate_to_datetime

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(APP_DIR, "quick_copy_data.json")
CONFIG_FILE = os.path.join(APP_DIR, "verify_config.json")

# ---------- 视觉主题 ----------
FONT_FAMILY = "Microsoft YaHei UI"
FONT_MAIN = (FONT_FAMILY, 9)
FONT_BOLD = (FONT_FAMILY, 9, "bold")
FONT_KEY_BOLD = (FONT_FAMILY, 9, "bold")
FONT_SMALL = (FONT_FAMILY, 8)
FONT_TITLE = (FONT_FAMILY, 10, "bold")
FONT_CODE = ("Consolas", 18, "bold")

C_BG = "#eef1f6"        # 窗口底色
C_CARD = "#ffffff"      # 卡片
C_BORDER = "#e1e5ee"    # 细边框
C_ACCENT = "#3d63d6"    # 主色
C_ACCENT_D = "#2f4fb0"  # 主色深
C_TEXT = "#1f2733"      # 主文字
C_MUTED = "#8a93a5"     # 次文字
C_HOVER = "#eef3ff"     # 行 hover

WIDTH = 310
ROW_H = 34
HEADER_H = 38
MAX_BODY_H = 330

EMAIL_HOSTS = {
    "163.com": "imap.163.com", "126.com": "imap.126.com",
    "qq.com": "imap.qq.com", "foxmail.com": "imap.qq.com",
    "gmail.com": "imap.gmail.com",
    "outlook.com": "outlook.office365.com",
    "hotmail.com": "outlook.office365.com", "sina.com": "imap.sina.com",
}

CODE_NEAR_RE = re.compile(
    r"(?:验证码|校验码|校验代码|验证代码|动态码|verification code|security code|confirm code|otp|code)"
    r"[^\dA-Za-z]{0,12}(\d{4,8})", re.IGNORECASE)
NUM_RE = re.compile(r"(?<!\d)(\d{4,8})(?!\d)")

GWL_EXSTYLE = -20
WS_EX_COMPOSITED = 0x02000000


def double_buffered(widget):
    try:
        hwnd = widget.winfo_id()
        u32 = ctypes.windll.user32
        cur = u32.GetWindowLongPtrW(hwnd, GWL_EXSTYLE)
        u32.SetWindowLongPtrW(hwnd, GWL_EXSTYLE, cur | WS_EX_COMPOSITED)
    except Exception:
        pass


# ---------- Windows DPAPI ----------
class _Blob(ctypes.Structure):
    _fields_ = [("cbData", ctypes.wintypes.DWORD),
                ("pbData", ctypes.POINTER(ctypes.c_char))]


def dpapi_protect(text: str) -> str:
    raw = text.encode("utf-8")
    inb = _Blob(len(raw), ctypes.cast(ctypes.c_char_p(raw),
                                      ctypes.POINTER(ctypes.c_char)))
    out = _Blob()
    if not ctypes.windll.crypt32.CryptProtectData(
            ctypes.byref(inb), None, None, None, None, 0, ctypes.byref(out)):
        raise ctypes.WinError()
    blob = ctypes.string_at(out.pbData, out.cbData)
    ctypes.windll.kernel32.LocalFree(out.pbData)
    return base64.b64encode(blob).decode()


def dpapi_unprotect(b64: str) -> str:
    blob = base64.b64decode(b64)
    inb = _Blob(len(blob), ctypes.cast(ctypes.c_char_p(blob),
                                       ctypes.POINTER(ctypes.c_char)))
    out = _Blob()
    if not ctypes.windll.crypt32.CryptUnprotectData(
            ctypes.byref(inb), None, None, None, None, 0, ctypes.byref(out)):
        raise ctypes.WinError()
    raw = ctypes.string_at(out.pbData, out.cbData)
    ctypes.windll.kernel32.LocalFree(out.pbData)
    return raw.decode("utf-8")


# ---------- 邮件 ----------
def _decoded_str(raw):
    if raw is None:
        return ""
    out = []
    for txt, enc in decode_header(raw):
        if isinstance(txt, bytes):
            out.append(txt.decode(enc or "utf-8", errors="replace"))
        else:
            out.append(txt)
    return "".join(out)


def _msg_text(msg):
    chunks = []
    parts = msg.walk() if msg.is_multipart() else [msg]
    for part in parts:
        if msg.is_multipart() and part.get_content_type() not in (
                "text/plain", "text/html"):
            continue
        payload = part.get_payload(decode=True)
        if payload:
            chunks.append(payload.decode(part.get_content_charset() or "utf-8",
                                         errors="replace"))
    return re.sub(r"<[^>]+>", " ", "\n".join(chunks))


def extract_code(text: str):
    m = CODE_NEAR_RE.search(text)
    if m:
        return m.group(1)
    cands = NUM_RE.findall(text)
    return cands[0] if cands else None


@contextmanager
def ipv4_only():
    orig = socket.getaddrinfo

    def filtered(host, port=0, family=0, type_=0, proto=0, flags=0):
        return [i for i in orig(host, port, family, type_, proto, flags)
                if i[0] == socket.AF_INET]
    socket.getaddrinfo = filtered
    try:
        yield
    finally:
        socket.getaddrinfo = orig


def _netease_login(M, cfg):
    M.login(cfg["addr"], dpapi_unprotect(cfg["secret"]))
    if any(k in cfg["host"] for k in ("163.com", "126.com", "netease")):
        payload = '("name" "QuickDesk" "version" "2.0" "vendor" "local")'
        typ, data = M._simple_command("ID", payload)
        M._untagged_response(typ, data, "ID")


def _select_inbox(M):
    typ, data = M.select("INBOX", readonly=True)
    if typ != "OK":
        raise RuntimeError(f"打开收件箱失败：{data}")


def imap_baseline(cfg):
    with ipv4_only():
        M = imaplib.IMAP4_SSL(cfg["host"], 993, timeout=12)
    _netease_login(M, cfg)
    _select_inbox(M)
    typ, data = M.uid("search", "ALL")
    if typ != "OK":
        raise RuntimeError(f"SEARCH 失败：{data}")
    uids = [int(u) for u in data[0].split() if u.isdigit()]
    M.logout()
    return max(uids) if uids else 0


def imap_fetch_code(cfg):
    base = int(cfg.get("base_uid", 0))
    with ipv4_only():
        M = imaplib.IMAP4_SSL(cfg["host"], 993, timeout=12)
    _netease_login(M, cfg)
    _select_inbox(M)
    typ, data = M.uid("search", "UID", f"{base + 1}:*")
    if typ != "OK":
        M.logout()
        raise RuntimeError(f"SEARCH 失败：{data}")
    uids = [u for u in data[0].split() if u.isdigit()]
    for uid in reversed(uids):
        if int(uid) <= base:
            continue
        typ, msgdata = M.uid("fetch", uid, "(RFC822)")
        if not msgdata or not msgdata[0]:
            continue
        msg = email.message_from_bytes(msgdata[0][1])
        subject = _decoded_str(msg.get("Subject"))
        code = extract_code(subject + "\n" + _msg_text(msg))
        if code:
            try:
                ts = parsedate_to_datetime(msg.get("Date")).timestamp()
            except Exception:
                ts = time.time()
            M.logout()
            return int(uid), code, ts
    M.logout()
    return None


# ---------- ADB / 短信 ----------
def find_adb():
    p = shutil.which("adb")
    if p:
        return p
    local = os.path.join(APP_DIR, "adb", "adb.exe")
    return local if os.path.exists(local) else None


def adb_devices():
    adb = find_adb()
    if not adb:
        return None, []
    out = subprocess.run([adb, "devices"], capture_output=True, timeout=10)
    devs = []
    for line in out.stdout.decode("utf-8", errors="replace").splitlines()[1:]:
        parts = line.strip().split("\t")
        if len(parts) == 2 and parts[1] == "device":
            devs.append(parts[0])
    return adb, devs


def sms_fetch_code(cfg):
    adb = find_adb()
    cmd = [adb]
    if cfg.get("serial"):
        cmd += ["-s", cfg["serial"]]
    cmd += ["shell", "content", "query", "--uri", "content://sms/inbox",
            "--projection", "date,body", "--sort", "date DESC LIMIT 15"]
    out = subprocess.run(cmd, capture_output=True, timeout=12)
    base = int(cfg.get("base_date", 0))
    for line in out.stdout.decode("utf-8", errors="replace").splitlines():
        m = re.match(r"Row:\s*\d+\s*date=(\d+)\s*body=(.*)", line, re.S)
        if not m:
            continue
        date_ms, body = int(m.group(1)), m.group(2).rstrip(",")
        if date_ms <= base:
            continue
        code = extract_code(body)
        if code:
            return date_ms, code
    return None


def install_adb(progress_cb):
    urls = [
        "https://dl.google.com/android/repository/platform-tools-latest-windows.zip",
        "https://mirrors.bfsu.edu.cn/android/repository/platform-tools-latest-windows.zip",
        "https://mirrors.tuna.tsinghua.edu.cn/android/repository/platform-tools-latest-windows.zip",
    ]
    tmp_zip = os.path.join(APP_DIR, "platform-tools.zip")
    last = None
    for url in urls:
        try:
            progress_cb(f"下载中：{url.split('/')[2]}…")
            req = urllib.request.Request(url,
                                         headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=60) as r, \
                    open(tmp_zip, "wb") as f:
                f.write(r.read())
            break
        except Exception as e:
            last = e
    else:
        raise RuntimeError(f"所有源失败：{last}")
    progress_cb("解压中…")
    with zipfile.ZipFile(tmp_zip) as z:
        z.extractall(APP_DIR)
    dst = os.path.join(APP_DIR, "adb")
    if os.path.exists(dst):
        shutil.rmtree(dst)
    os.rename(os.path.join(APP_DIR, "platform-tools"), dst)
    os.remove(tmp_zip)
    progress_cb("adb 安装完成")


# ================= 主应用 =================
class QuickCopyApp:
    def __init__(self, root):
        self.root = root
        self.groups = []
        self.active_group = 0
        self.pos = None
        self.collapsed = False
        self.pin_on = True
        self.masked = False
        self.verify_cfg = {"show_panel": True, "email": None, "sms": None}

        self.editor_visible = False
        self.geditor_visible = False
        self.editing_index = None
        self._drag = None
        self._delete_arm = None
        self._status_after = None
        self._row_widgets = {}
        self._watch_job = None
        self._watch_left = 0
        self._smooth_target = 0.0
        self._smooth_job = None
        self._hover_locked = False
        self._last_pix = None
        self._content_px = 1

        self.load_data()
        self.load_verify_cfg()

        root.overrideredirect(True)
        root.attributes("-topmost", self.pin_on)
        root.config(bg=C_BG)

        self.build_header()
        self.body = tk.Frame(root, bg=C_BG)
        self.build_groupbar()
        self.build_vpanel()
        self.build_geditor()
        self.build_editor()
        self.build_list()
        self.build_footer()

        self.header.pack(fill="x")
        self.body.pack(fill="both", expand=True)
        self.groupbar.pack(fill="x", padx=8, pady=(7, 2))
        if self.verify_cfg.get("show_panel", True):
            self.vpanel.pack(fill="x", padx=8, pady=(4, 0))
        self.list_wrap.pack(fill="both", expand=True, padx=4, pady=(4, 2))
        self.footer.pack(fill="x", padx=6, pady=(0, 4))

        self.refresh_group_ui()
        self.refresh_rows()

        x, y = self.pos or (70, 110)
        if self.collapsed:
            self.body.pack_forget()
            self.btn_collapse.config(text="▲")
        root.geometry(f"{WIDTH}x{self.current_height()}+{x}+{y}")

        def warm():
            adb = find_adb()
            if adb:
                subprocess.run([adb, "start-server"], capture_output=True)
        threading.Thread(target=warm, daemon=True).start()

    # ---------- 持久化 ----------
    def load_data(self):
        d = None
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, "r", encoding="utf-8") as f:
                    d = json.load(f)
            except Exception:
                d = None
        if d and d.get("groups"):
            self.groups = d["groups"]
            self.active_group = min(d.get("active_group", 0),
                                    len(self.groups) - 1)
            self.pos = d.get("pos")
            self.collapsed = d.get("collapsed", False)
            self.pin_on = d.get("pin_on", True)
            self.masked = d.get("masked", False)
        elif d and d.get("fields"):
            # 旧版平铺数据迁移为分组
            self.groups = [
                {"name": "求职信息", "fields": d.get("fields", [])},
                {"name": "应用密码", "fields": []},
            ]
            self.pos = d.get("pos")
            self.collapsed = d.get("collapsed", False)
            self.pin_on = d.get("pin_on", True)
            self.masked = d.get("masked", False)
        else:
            self.groups = [
                {"name": "求职信息", "fields": []},
                {"name": "应用密码", "fields": []},
            ]

    def save_data(self):
        try:
            with open(DATA_FILE, "w", encoding="utf-8") as f:
                json.dump({"groups": self.groups,
                           "active_group": self.active_group,
                           "pos": self.pos, "collapsed": self.collapsed,
                           "pin_on": self.pin_on, "masked": self.masked},
                          f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def load_verify_cfg(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    self.verify_cfg.update(json.load(f))
            except Exception:
                pass

    def save_verify_cfg(self):
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.verify_cfg, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    @property
    def fields(self):
        return self.groups[self.active_group]["fields"]

    # ---------- 头部 ----------
    def hbtn(self, text, cmd):
        b = tk.Label(self.header, text=text, bg=C_ACCENT, fg="#ffffff",
                     font=FONT_MAIN, padx=7, cursor="hand2")
        b.bind("<Button-1>", lambda e: cmd())
        b.bind("<Enter>", lambda e: b.config(bg=C_ACCENT_D))
        b.bind("<Leave>", lambda e: b.config(bg=C_ACCENT))
        return b

    def build_header(self):
        self.header = tk.Frame(self.root, bg=C_ACCENT, height=HEADER_H)
        self.header.pack_propagate(False)
        title = tk.Label(self.header, text="  快捷工作台", bg=C_ACCENT,
                         fg="#ffffff", font=FONT_TITLE)
        title.pack(side="left")
        for w in (self.header, title):
            w.bind("<Button-1>", self.start_drag)
            w.bind("<B1-Motion>", self.on_drag)
            w.bind("<ButtonRelease-1>", self.end_drag)

        self.btn_close = self.hbtn("✕", self.close_app)
        self.btn_collapse = self.hbtn("▼", self.toggle_collapse)
        self.btn_pin = self.hbtn("📌", self.toggle_pin)
        self.btn_code = self.hbtn("码", self.toggle_vpanel)
        self.btn_mask = self.hbtn("隐", self.toggle_mask)
        self.btn_add = self.hbtn("＋", self.click_add)
        for b in (self.btn_close, self.btn_collapse, self.btn_pin,
                  self.btn_code, self.btn_mask, self.btn_add):
            b.pack(side="right")
        if not self.pin_on:
            self.btn_pin.config(fg="#b9c6ee")
        self.btn_mask.config(text="显" if self.masked else "隐")

    # ---------- 分组栏 ----------
    def build_groupbar(self):
        self.groupbar = tk.Frame(self.body, bg=C_BG)
        self.group_var = tk.StringVar()
        ttk.Style().configure("G.TCombobox", padding=2)
        self.group_cb = ttk.Combobox(
            self.groupbar, textvariable=self.group_var, state="readonly",
            font=FONT_SMALL, width=16, style="G.TCombobox")
        self.group_cb.pack(side="left")
        self.group_cb.bind("<<ComboboxSelected>>", self.on_pick_group)
        self.gbtn("＋", self.show_geditor)
        self.gbtn("✎", self.rename_group)
        self.gbtn("✕", self.delete_group)

    def gbtn(self, text, cmd):
        b = tk.Label(self.groupbar, text=text, bg=C_BG, fg=C_ACCENT,
                     font=FONT_BOLD, padx=4, cursor="hand2")
        b.pack(side="left", padx=1)
        b.bind("<Button-1>", lambda e: cmd())

    def refresh_group_ui(self):
        names = [g["name"] for g in self.groups]
        self.group_cb["values"] = names
        self.group_var.set(names[self.active_group])

    def on_pick_group(self, _):
        name = self.group_var.get()
        for i, g in enumerate(self.groups):
            if g["name"] == name:
                self.active_group = i
                break
        self.refresh_rows()
        self.save_data()

    def build_geditor(self):
        self.geditor = tk.Frame(self.body, bg=C_CARD, bd=1, relief="solid")
        self.gname_var = tk.StringVar()
        e = tk.Entry(self.geditor, textvariable=self.gname_var,
                     font=FONT_SMALL, width=15, relief="flat")
        ok = tk.Label(self.geditor, text="确定", bg=C_ACCENT, fg="#fff",
                      font=FONT_SMALL, padx=7, cursor="hand2")
        no = tk.Label(self.geditor, text="取消", bg="#eef0f4", fg="#666",
                      font=FONT_SMALL, padx=7, cursor="hand2")
        e.pack(side="left", padx=(6, 4), pady=4)
        ok.pack(side="left", padx=2); no.pack(side="left", padx=2)
        ok.bind("<Button-1>", lambda e: self.commit_geditor())
        no.bind("<Button-1>", lambda e: self.hide_geditor())
        e.bind("<Return>", lambda e: self.commit_geditor())
        e.bind("<Escape>", lambda e: self.hide_geditor())

    def show_geditor(self):
        if not self.geditor_visible:
            self.geditor_visible = True
            self.gname_var.set("")
            self.geditor.pack(fill="x", padx=8, pady=(2, 0),
                              before=self.list_wrap)
            self.relayout()

    def hide_geditor(self):
        self.geditor_visible = False
        self.geditor.pack_forget()
        self.relayout()

    def commit_geditor(self):
        name = self.gname_var.get().strip()
        if not name:
            return
        if any(g["name"] == name for g in self.groups):
            self.flash_status(f"组名“{name}”已存在", warn=True)
            return
        self.groups.append({"name": name, "fields": []})
        self.active_group = len(self.groups) - 1
        self.hide_geditor()
        self.refresh_group_ui()
        self.refresh_rows()
        self.save_data()

    def rename_group(self):
        cur = self.groups[self.active_group]
        win = tk.Toplevel(self.root)
        win.overrideredirect(True)
        win.config(bg=C_BORDER)
        var = tk.StringVar(value=cur["name"])
        e = tk.Entry(win, textvariable=var, font=FONT_MAIN, width=16,
                     relief="flat")
        e.pack(padx=1, pady=1)
        e.focus_set(); e.select_range(0, "end")

        def save(_=None):
            nm = var.get().strip()
            if nm and not any(i != self.active_group
                              and self.groups[i]["name"] == nm
                              for i in range(len(self.groups))):
                cur["name"] = nm
            win.destroy()
            self.refresh_group_ui()
            self.save_data()
        e.bind("<Return>", save)
        e.bind("<Escape>", lambda e: win.destroy())
        win.geometry(f"+{self.root.winfo_x()+30}+{self.root.winfo_y()+50}")

    def delete_group(self):
        if len(self.groups) <= 1:
            self.flash_status("至少保留一个组，不能删除", warn=True)
            return
        cur_name = self.groups[self.active_group]["name"]
        win = tk.Toplevel(self.root)
        win.overrideredirect(True)
        f = tk.Frame(win, bg=C_CARD, bd=1, relief="solid")
        f.pack()
        tk.Label(f, text=f"删除整组“{cur_name}”？\n组内条目会一起删除",
                 bg=C_CARD, fg=C_TEXT, font=FONT_SMALL, justify="left"
                 ).pack(padx=10, pady=8)
        bs = tk.Frame(f, bg=C_CARD)
        bs.pack(pady=(0, 8))

        def yes():
            win.destroy()
            del self.groups[self.active_group]
            self.active_group = max(0, self.active_group - 1)
            self.refresh_group_ui()
            self.refresh_rows()
            self.save_data()
        for t, c, fn in (("删除", "#e53935", yes),
                         ("取消", "#888", win.destroy)):
            b = tk.Label(bs, text=t, bg=c, fg="#fff", font=FONT_SMALL,
                         padx=10, pady=2, cursor="hand2")
            b.pack(side="left", padx=3)
        win.geometry(f"+{self.root.winfo_x()+20}+{self.root.winfo_y()+60}")

    # ---------- 验证码面板 ----------
    def build_vpanel(self):
        self.vpanel = tk.Frame(self.body, bg=C_CARD, bd=1, relief="solid")
        tk.Label(self.vpanel, text="验证码助手", bg=C_CARD, fg=C_ACCENT,
                 font=FONT_BOLD, anchor="w").pack(fill="x", padx=8,
                                                  pady=(6, 2))
        self.code_btn = tk.Label(self.vpanel, text="———", font=FONT_CODE,
                                 bg="#f7f9fd", fg=C_ACCENT, relief="solid",
                                 bd=1, height=1, cursor="hand2")
        self.code_btn.pack(fill="x", padx=8)
        self.code_btn.bind("<Button-1>", lambda e: self.copy_current_code())
        bs = tk.Frame(self.vpanel, bg=C_CARD)
        bs.pack(fill="x", padx=6, pady=5)
        for t, fn in (("邮箱", self.fetch_email_code),
                      ("短信", self.fetch_sms_code),
                      ("绑定", self.open_bind_dialog),
                      ("停止", self.stop_watch)):
            b = tk.Label(bs, text=t, bg=C_ACCENT, fg="#fff", font=FONT_SMALL,
                         padx=7, pady=2, cursor="hand2")
            b.pack(side="left", padx=2)
            b.bind("<Enter>", lambda e, w=b: w.config(bg=C_ACCENT_D))
            b.bind("<Leave>", lambda e, w=b: w.config(bg=C_ACCENT))
        self.code_status = tk.Label(self.vpanel, text="未获取", bg=C_CARD,
                                    fg=C_MUTED, font=FONT_SMALL, anchor="w")
        self.code_status.pack(fill="x", padx=10, pady=(0, 5))
        self._current_code = ""

    def toggle_vpanel(self):
        if self.vpanel.winfo_ismapped():
            self.vpanel.pack_forget()
            self.verify_cfg["show_panel"] = False
            self.stop_watch()
        else:
            self.vpanel.pack(fill="x", padx=8, pady=(4, 0),
                             before=self.list_wrap)
            self.verify_cfg["show_panel"] = True
        self.save_verify_cfg()
        self.relayout()

    def set_code_status(self, text, color=C_MUTED):
        self.code_status.config(text=text, fg=color)

    def copy_current_code(self):
        if not self._current_code:
            self.set_code_status("还没有验证码", "#e53935")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(self._current_code)
        self.root.update()
        self.set_code_status(f"已复制 {self._current_code}", C_ACCENT)

    def on_code_found(self, code, source, ts, advance):
        self._current_code = code
        self.code_btn.config(text=code)
        self.stop_watch()
        self.root.clipboard_clear()
        self.root.clipboard_append(code)
        self.root.update()
        if source == "email":
            self.verify_cfg["email"]["base_uid"] = advance
        else:
            self.verify_cfg["sms"]["base_date"] = advance
        self.save_verify_cfg()
        ago = max(0, int(time.time() - ts))
        tag = "邮箱" if source == "email" else "短信"
        self.set_code_status(f"{tag}验证码已获取并自动复制（{ago}秒前）",
                             C_ACCENT)

    # 邮箱
    def fetch_email_code(self):
        cfg = self.verify_cfg.get("email")
        if not cfg:
            self.set_code_status("请先点“绑定”绑定邮箱", "#e53935")
            self.open_bind_dialog()
            return
        self.set_code_status("查询邮箱中…")
        self.start_watch("email")
        threading.Thread(target=self.email_worker, args=(cfg,),
                         daemon=True).start()

    def email_worker(self, cfg):
        try:
            found = imap_fetch_code(cfg)
            if found:
                uid, code, ts = found
                self.root.after(0, lambda: self.on_code_found(
                    code, "email", ts, uid))
            else:
                self.root.after(0, self.watch_no_code)
        except Exception as ex:
            self.root.after(0, lambda: self.watch_fail("邮箱", str(ex)))

    # 短信
    def fetch_sms_code(self):
        cfg = self.verify_cfg.get("sms")
        if not cfg:
            self.set_code_status("请先绑定安卓手机", "#e53935")
            self.open_bind_dialog()
            return
        _, devs = adb_devices()
        if not devs:
            self.set_code_status("手机未连接", "#e53935")
            return
        self.set_code_status("读取短信中…")
        self.start_watch("sms")
        threading.Thread(target=self.sms_worker, args=(cfg,),
                         daemon=True).start()

    def sms_worker(self, cfg):
        try:
            found = sms_fetch_code(cfg)
            if found:
                dms, code = found
                self.root.after(0, lambda: self.on_code_found(
                    code, "sms", dms / 1000, dms))
            else:
                self.root.after(0, self.watch_no_code)
        except Exception as ex:
            self.root.after(0, lambda: self.watch_fail("短信", str(ex)))

    def start_watch(self, source):
        self.stop_watch()
        self._watch_source = source
        self._watch_left = 25

    def watch_no_code(self):
        if self._watch_left <= 0:
            return
        self._watch_left -= 1
        if self._watch_left <= 0:
            self.set_code_status("等待超时：未收到新验证码", "#e53935")
            return
        self.set_code_status(f"等待新验证码…（{self._watch_left*3}s）")
        self._watch_job = self.root.after(3000, self.watch_retry)

    def watch_retry(self):
        self._watch_job = None
        if self._watch_source == "email":
            cfg = self.verify_cfg.get("email")
            if cfg:
                threading.Thread(target=self.email_worker, args=(cfg,),
                                 daemon=True).start()
        else:
            cfg = self.verify_cfg.get("sms")
            if cfg:
                threading.Thread(target=self.sms_worker, args=(cfg,),
                                 daemon=True).start()

    def watch_fail(self, tag, msg):
        self.stop_watch()
        self.set_code_status(f"{tag}失败：{msg[:40]}", "#e53935")

    def stop_watch(self):
        if self._watch_job:
            self.root.after_cancel(self._watch_job)
            self._watch_job = None
        self._watch_left = 0

    # ---------- 字段编辑器 ----------
    def build_editor(self):
        self.editor = tk.Frame(self.body, bg=C_CARD, bd=1, relief="solid")
        self.name_var = tk.StringVar()
        self.value_var = tk.StringVar()
        self.bold_var = tk.BooleanVar(value=False)
        self.pick_color = "#666666"

        e1 = tk.Entry(self.editor, textvariable=self.name_var, font=FONT_MAIN,
                      relief="flat", width=9)
        e2 = tk.Entry(self.editor, textvariable=self.value_var, font=FONT_MAIN,
                      relief="flat")
        ok = tk.Label(self.editor, text="确定", bg=C_ACCENT, fg="#fff",
                      font=FONT_SMALL, padx=8, cursor="hand2")
        no = tk.Label(self.editor, text="取消", bg="#eef0f4", fg="#555",
                      font=FONT_SMALL, padx=8, cursor="hand2")
        ok.bind("<Button-1>", lambda e: self.commit_editor())
        no.bind("<Button-1>", lambda e: self.hide_editor())
        e1.grid(row=0, column=0, padx=(6, 3), pady=(5, 2), sticky="w")
        e2.grid(row=0, column=1, pady=(5, 2), sticky="ew")
        ok.grid(row=0, column=2, padx=(6, 2), pady=(5, 2))
        no.grid(row=0, column=3, padx=(2, 6), pady=(5, 2))
        self.editor.grid_columnconfigure(1, weight=1)
        for e in (e1, e2):
            e.bind("<Return>", lambda ev: self.commit_editor())
            e.bind("<Escape>", lambda ev: self.hide_editor())
        self.editor_entries = (e1, e2)

        sb = tk.Frame(self.editor, bg=C_CARD)
        sb.grid(row=1, column=0, columnspan=4, sticky="ew", padx=4,
                pady=(0, 5))
        self.bold_btn = tk.Label(sb, text="B", bg="#fff", fg="#333",
                                 font=FONT_KEY_BOLD, padx=7, cursor="hand2",
                                 bd=1, relief="solid")
        self.bold_btn.pack(side="left")
        self.bold_btn.bind("<Button-1>", lambda e: self.toggle_bold_pick())
        tk.Label(sb, text="key颜色", bg=C_CARD, fg=C_MUTED, font=FONT_SMALL
                 ).pack(side="left", padx=(8, 3))
        self.swatches = {}
        for c in ("#333333", "#e53935", C_ACCENT, "#2e9b46",
                  "#f08c1a", "#8e44ad"):
            s = tk.Label(sb, text="  ", bg=c, cursor="hand2", bd=1,
                         relief="solid")
            s.pack(side="left", padx=1)
            s.bind("<Button-1>", lambda e, col=c: self.pick_key_color(col))
            self.swatches[c] = s

    def toggle_bold_pick(self):
        on = not self.bold_var.get()
        self.bold_var.set(on)
        self.bold_btn.config(bg="#333" if on else "#fff",
                             fg="#fff" if on else "#333")

    def pick_key_color(self, color):
        self.pick_color = color
        for c, s in self.swatches.items():
            s.config(relief="sunken" if c == color else "solid")

    def reset_style_pick(self):
        self.bold_var.set(False)
        self.bold_btn.config(bg="#fff", fg="#333")
        self.pick_color = "#666666"
        for s in self.swatches.values():
            s.config(relief="solid")

    def click_add(self):
        if self.editor_visible and self.editing_index is None:
            self.hide_editor()
            return
        self.editing_index = None
        self.name_var.set(""); self.value_var.set("")
        self.reset_style_pick()
        self.show_editor()
        self.editor_entries[0].focus_set()

    def click_edit(self, idx):
        item = self.fields[idx]
        self.editing_index = idx
        self.name_var.set(item["name"])
        self.value_var.set(item["value"])
        self.bold_var.set(bool(item.get("bold")))
        self.bold_btn.config(bg="#333" if item.get("bold") else "#fff",
                             fg="#fff" if item.get("bold") else "#333")
        self.pick_color = item.get("color", "#666666")
        for c, s in self.swatches.items():
            s.config(relief="sunken" if c == self.pick_color else "solid")
        self.show_editor()
        self.editor_entries[1].focus_set()
        self.editor_entries[1].select_range(0, "end")

    def show_editor(self):
        if not self.editor_visible:
            self.editor_visible = True
            self.editor.pack(fill="x", padx=8, pady=(4, 0),
                             before=self.list_wrap)
        self.relayout()

    def hide_editor(self):
        self.editor_visible = False
        self.editing_index = None
        self.editor.pack_forget()
        self.relayout()

    def commit_editor(self):
        name = self.name_var.get().strip() or "未命名"
        for i, f in enumerate(self.fields):
            if i != self.editing_index and f["name"] == name:
                self.flash_status(f"本组内“{name}”已存在，必须唯一",
                                  warn=True)
                return
        value = self.value_var.get().strip()
        item = {"name": name, "value": value, "color": self.pick_color,
                "bold": self.bold_var.get()}
        if self.editing_index is None:
            self.fields.append(item)
        else:
            self.fields[self.editing_index] = item
        self.hide_editor()
        self.refresh_rows()
        self.save_data()

    # ---------- 列表 ----------
    def build_list(self):
        self.list_wrap = tk.Frame(self.body, bg=C_BG)
        self.canvas = tk.Canvas(self.list_wrap, bg=C_BG, highlightthickness=0)
        sb = tk.Scrollbar(self.list_wrap, orient="vertical",
                          command=self.canvas.yview)
        self.canvas.config(yscrollcommand=sb.set)
        self.inner = tk.Frame(self.canvas, bg=C_BG)
        self.inner_id = self.canvas.create_window((0, 0), window=self.inner,
                                                  anchor="nw")
        self.canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        def cfg(e):
            self._content_px = e.height
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.inner.bind("<Configure>", cfg)
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(
            self.inner_id, width=e.width))
        self.canvas.bind("<Enter>", lambda e: self.root.bind_all(
            "<MouseWheel>", self.on_wheel))
        self.canvas.bind("<Leave>", lambda e: self.root.unbind_all(
            "<MouseWheel>"))
        double_buffered(self.canvas)
        double_buffered(self.inner)

    def refresh_rows(self):
        for w in self.inner.winfo_children():
            w.destroy()
        self._row_widgets.clear()
        if not self.fields:
            tk.Label(self.inner, text="本组还没有条目\n点右上角 ＋ 添加",
                     bg=C_BG, fg=C_MUTED, font=FONT_MAIN, pady=18,
                     justify="center").pack(fill="x")
            self.relayout()
            return
        for i, item in enumerate(self.fields):
            card = tk.Frame(self.inner, bg=C_CARD, height=ROW_H, bd=1,
                            relief="solid", highlightbackground=C_BORDER)
            card.pack(fill="x", padx=5, pady=2)
            card.pack_propagate(False)
            double_buffered(card)

            name = tk.Label(card, text=item["name"], bg=C_CARD,
                            fg=item.get("color", "#666666"),
                            font=FONT_KEY_BOLD if item.get("bold")
                            else FONT_MAIN, width=11, anchor="e")
            if self.masked:
                disp = "••••••" if item["value"] else "（空）"
            else:
                v = item["value"]
                disp = v if len(v) <= 17 else v[:17] + "…"
                if not disp:
                    disp = "（空，点✎填写）"
            value = tk.Label(card, text=disp, bg=C_CARD, fg=C_TEXT,
                             font=FONT_MAIN, anchor="w")
            edit = tk.Label(card, text="✎", bg=C_CARD, fg=C_MUTED,
                            font=FONT_MAIN, padx=4, cursor="hand2")
            dele = tk.Label(card, text="✕", bg=C_CARD, fg=C_MUTED,
                            font=FONT_MAIN, padx=4, cursor="hand2")
            edit.bind("<Button-1>", lambda e, x=i: self.click_edit(x))
            dele.bind("<Button-1>", lambda e, x=i, b=dele:
                      self.click_delete(x, b))
            name.grid(row=0, column=0, padx=(4, 4), sticky="ns")
            value.grid(row=0, column=1, sticky="nsew")
            edit.grid(row=0, column=2)
            dele.grid(row=0, column=3, padx=(0, 4))
            card.grid_columnconfigure(1, weight=1)
            for w in (card, name, value):
                w.bind("<Button-1>", lambda e, x=i: self.do_copy(x))
                w.bind("<Enter>", lambda e, x=i: self.hover(x, True))
                w.bind("<Leave>", lambda e, x=i: self.hover(x, False))
            self._row_widgets[i] = [card, name, value]
        self.relayout()

    def hover(self, idx, enter):
        if self._hover_locked:
            return
        c = C_HOVER if enter else C_CARD
        for w in self._row_widgets.get(idx, []):
            w.config(bg=c)

    def reset_all_hover(self):
        for ws in self._row_widgets.values():
            for w in ws:
                w.config(bg=C_CARD)

    def do_copy(self, idx):
        text = self.fields[idx]["value"]
        if not text:
            self.flash_status("该条目为空，点 ✎ 填写", warn=True)
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update()
        self.flash_status(f"已复制【{self.fields[idx]['name']}】")

    def click_delete(self, idx, btn):
        if self._delete_arm == idx:
            del self.fields[idx]
            self._delete_arm = None
            self.refresh_rows()
            self.save_data()
            return
        self._delete_arm = idx
        btn.config(text="确认?", fg="#e53935")

        def back():
            if self._delete_arm == idx:
                self._delete_arm = None
                btn.config(text="✕", fg=C_MUTED)
        self.root.after(2000, back)

    def build_footer(self):
        self.footer = tk.Label(
            self.body, text="点条目复制 · 数据本地保存", bg=C_BG, fg=C_MUTED,
            font=FONT_SMALL, anchor="w", padx=10)

    def flash_status(self, text, warn=False):
        if self._status_after:
            self.root.after_cancel(self._status_after)
        self.footer.config(text=text, fg="#e53935" if warn else C_ACCENT)
        self._status_after = self.root.after(
            1600, lambda: self.footer.config(
                text="点条目复制 · 数据本地保存", fg=C_MUTED))

    # ---------- 滚动 ----------
    def on_wheel(self, e):
        starting = self._smooth_job is None
        if starting:
            self._smooth_target = float(self.canvas.yview()[0])
            self._last_pix = None
            self.reset_all_hover()
            self._hover_locked = True
        self._smooth_target = max(
            0.0, min(1.0,
                      self._smooth_target + (e.delta / 120.0) * 0.075))
        if starting:
            self.smooth_tick()

    def smooth_tick(self):
        cur = float(self.canvas.yview()[0])
        diff = self._smooth_target - cur
        if abs(diff) < 0.0005:
            self.canvas.yview_moveto(self._smooth_target)
            self._smooth_job = None
            self._hover_locked = False
            return
        nxt = cur + diff * 0.25
        view_h = self.canvas.winfo_height()
        total = max(1, self._content_px - view_h)
        pix = int(nxt * total)
        if pix != self._last_pix:
            self.canvas.yview_moveto(nxt)
            self._last_pix = pix
        self._smooth_job = self.root.after(10, self.smooth_tick)

    # ---------- 窗口行为 ----------
    def current_height(self):
        h = HEADER_H
        if not self.collapsed:
            n = max(len(self.fields), 1)
            h += min(n * (ROW_H + 4) + 10, MAX_BODY_H)
            if self.editor_visible:
                h += 64
            if self.geditor_visible:
                h += 32
            if self.verify_cfg.get("show_panel", True):
                h += 118
            h += 22
        return h

    def relayout(self):
        self.root.update_idletasks()
        x, y = self.root.winfo_x(), self.root.winfo_y()
        self.root.geometry(f"{WIDTH}x{self.current_height()}+{x}+{y}")

    def toggle_collapse(self):
        self.collapsed = not self.collapsed
        if self.collapsed:
            self.body.pack_forget()
            self.btn_collapse.config(text="▲")
        else:
            self.body.pack(fill="both", expand=True)
            self.btn_collapse.config(text="▼")
        self.relayout()
        self.save_data()

    def toggle_pin(self):
        self.pin_on = not self.pin_on
        self.root.attributes("-topmost", self.pin_on)
        self.btn_pin.config(fg="#fff" if self.pin_on else "#b9c6ee")
        self.save_data()

    def toggle_mask(self):
        self.masked = not self.masked
        self.btn_mask.config(text="显" if self.masked else "隐")
        self.refresh_rows()
        self.save_data()

    def start_drag(self, e):
        self._drag = (e.x_root - self.root.winfo_x(),
                      e.y_root - self.root.winfo_y())

    def on_drag(self, e):
        if self._drag:
            dx, dy = self._drag
            self.root.geometry(f"+{e.x_root-dx}+{e.y_root-dy}")

    def end_drag(self, e):
        self._drag = None
        self.pos = [self.root.winfo_x(), self.root.winfo_y()]
        self.save_data()

    def close_app(self):
        self.pos = [self.root.winfo_x(), self.root.winfo_y()]
        self.save_data()
        self.root.destroy()

    def open_bind_dialog(self):
        BindDialog(self)


# ================= 绑定对话框 =================
class BindDialog:
    def __init__(self, app):
        self.app = app
        win = tk.Toplevel(app.root)
        self.win = win
        win.title("绑定邮箱 / 手机")
        win.geometry("400x520")
        win.resizable(False, False)
        win.transient(app.root)
        win.grab_set()
        win.configure(bg="#f4f6fa")

        tk.Label(win, text="① 邮箱绑定（自动获取邮件验证码）",
                 bg="#f4f6fa", font=FONT_BOLD, anchor="w"
                 ).pack(fill="x", padx=12, pady=(10, 2))
        f1 = tk.Frame(win, bg="#f4f6fa"); f1.pack(fill="x", padx=12)
        self.addr_var = tk.StringVar()
        self.secret_var = tk.StringVar()
        self.host_var = tk.StringVar()
        old = app.verify_cfg.get("email")
        if old:
            self.addr_var.set(old.get("addr", ""))
            self.host_var.set(old.get("host", ""))
        self.row(f1, "邮箱地址", self.addr_var, 0)
        self.row(f1, "IMAP授权码", self.secret_var, 1, "*")
        self.row(f1, "IMAP服务器", self.host_var, 2)
        self.addr_var.trace_add("write", self.auto_host)
        tk.Label(win,
                 text="授权码不是登录密码：邮箱网页版 → 设置 →\n"
                      "POP3/SMTP/IMAP → 开启IMAP → 生成授权码",
                 bg="#f4f6fa", fg=C_MUTED, font=FONT_SMALL, justify="left",
                 anchor="w").pack(fill="x", padx=14)
        b = tk.Frame(win, bg="#f4f6fa"); b.pack(fill="x", padx=12, pady=4)
        self.dbtn(b, "保存并绑定邮箱", self.save_email)
        self.mail_st = tk.Label(win, text="", bg="#f4f6fa", fg=C_ACCENT,
                                font=FONT_SMALL, anchor="w")
        self.mail_st.pack(fill="x", padx=14)

        tk.Label(win, text="② 安卓手机绑定（短信验证码 / USB）",
                 bg="#f4f6fa", font=FONT_BOLD, anchor="w"
                 ).pack(fill="x", padx=12, pady=(8, 2))
        self.sms_st = tk.Label(win, text="", bg="#f4f6fa", fg=C_MUTED,
                               font=FONT_SMALL, anchor="w")
        self.sms_st.pack(fill="x", padx=14)
        f2 = tk.Frame(win, bg="#f4f6fa"); f2.pack(fill="x", padx=12, pady=2)
        self.pair_addr = tk.StringVar()
        self.pair_code = tk.StringVar()
        self.conn_addr = tk.StringVar()
        self.row(f2, "无线配对 地址:端口", self.pair_addr, 0, width=15)
        self.row(f2, "配对码", self.pair_code, 1, width=15)
        self.row(f2, "连接 地址:端口", self.conn_addr, 2, width=15)
        sb = tk.Frame(win, bg="#f4f6fa"); sb.pack(fill="x", padx=12, pady=4)
        self.dbtn(sb, "检测设备", self.check_devices)
        self.dbtn(sb, "配对", self.do_pair)
        self.dbtn(sb, "连接", self.do_connect)
        self.dbtn(sb, "下载adb", self.download_adb)
        self.dbtn(sb, "绑定设备", self.bind_sms)
        tk.Label(win,
                 text="USB方式：开发者选项开USB调试→插线→允许调试→\n"
                      "点“检测设备”→点“绑定设备”即可",
                 bg="#f4f6fa", fg=C_MUTED, font=FONT_SMALL, justify="left",
                 anchor="w").pack(fill="x", padx=14, pady=(2, 0))
        self.dbtn(win, "关闭", win.destroy, "bottom")
        self.check_devices()

    def row(self, p, label, var, r, show=None, width=22):
        tk.Label(p, text=label, bg="#f4f6fa", font=FONT_SMALL, width=14,
                 anchor="e").grid(row=r, column=0, pady=2, sticky="e")
        tk.Entry(p, textvariable=var, font=FONT_SMALL, width=width,
                 show=show).grid(row=r, column=1, pady=2, sticky="w")

    def dbtn(self, p, text, cmd, side="left"):
        b = tk.Label(p, text=text, bg=C_ACCENT, fg="#fff", font=FONT_SMALL,
                     padx=8, pady=3, cursor="hand2")
        b.pack(side=side, padx=3, pady=3)
        b.bind("<Button-1>", lambda e: cmd())
        b.bind("<Enter>", lambda e: b.config(bg=C_ACCENT_D))
        b.bind("<Leave>", lambda e: b.config(bg=C_ACCENT))

    def set(self, w, t, c=C_ACCENT):
        w.config(text=t, fg=c)

    def auto_host(self, *a):
        addr = self.addr_var.get().strip()
        if "@" in addr:
            host = EMAIL_HOSTS.get(addr.split("@", 1)[1].lower())
            if host and not self.host_var.get():
                self.host_var.set(host)

    def save_email(self):
        addr = self.addr_var.get().strip()
        secret = self.secret_var.get().strip()
        host = self.host_var.get().strip()
        if not (addr and secret and host):
            self.set(self.mail_st, "三项都要填写", "#e53935")
            return
        cfg = {"addr": addr, "host": host,
               "secret": dpapi_protect(secret), "base_uid": 0}
        self.set(self.mail_st, "登录建立基线中…")

        def work():
            try:
                cfg["base_uid"] = imap_baseline(cfg)
                self.app.verify_cfg["email"] = cfg
                self.app.save_verify_cfg()
                self.win.after(0, lambda: self.set(
                    self.mail_st, f"绑定成功，基线UID={cfg['base_uid']}"))
            except Exception as ex:
                self.win.after(0, lambda: self.set(
                    self.mail_st, f"失败：{str(ex)[:45]}", "#e53935"))
        threading.Thread(target=work, daemon=True).start()

    def check_devices(self):
        self.set(self.sms_st, "检测中…")

        def work():
            adb, devs = adb_devices()

            def ui():
                if not adb:
                    self.set(self.sms_st, "未检测到adb，可下载", "#e53935")
                elif not devs:
                    self.set(self.sms_st, "adb就绪，暂无设备", "#e53935")
                else:
                    self.set(self.sms_st, "已连接：" + "、".join(devs))
            self.win.after(0, ui)
        threading.Thread(target=work, daemon=True).start()

    def adb_cmd(self, args, w):
        adb = find_adb()
        if not adb:
            self.set(w, "未装adb", "#e53935")
            return

        def work():
            try:
                out = subprocess.run([adb] + args, capture_output=True,
                                     timeout=30)
                t = (out.stdout.decode("utf-8", "replace") +
                     out.stderr.decode("utf-8", "replace")).strip()
                self.win.after(0, lambda: self.set(w, t[:80] or "完成"))
            except Exception as ex:
                self.win.after(0, lambda: self.set(w, str(ex)[:50], "#e53935"))
        threading.Thread(target=work, daemon=True).start()

    def do_pair(self):
        a, c = self.pair_addr.get().strip(), self.pair_code.get().strip()
        if a and c:
            self.adb_cmd(["pair", a, c], self.sms_st)

    def do_connect(self):
        a = self.conn_addr.get().strip() or self.pair_addr.get().strip()
        if a:
            self.adb_cmd(["connect", a], self.sms_st)

    def download_adb(self):
        if find_adb():
            self.set(self.sms_st, "adb已存在")
            return

        def work():
            try:
                install_adb(lambda m: self.win.after(
                    0, lambda: self.set(self.sms_st, m)))
            except Exception as ex:
                self.win.after(0, lambda: self.set(
                    self.sms_st, "失败：" + str(ex)[:45], "#e53935"))
        threading.Thread(target=work, daemon=True).start()

    def bind_sms(self):
        _, devs = adb_devices()
        if not devs:
            self.set(self.sms_st, "没有已连接设备", "#e53935")
            return
        self.app.verify_cfg["sms"] = {
            "serial": devs[0], "base_date": int(time.time() * 1000)}
        self.app.save_verify_cfg()
        self.set(self.sms_st, f"已绑定 {devs[0]}")


def main():
    root = tk.Tk()
    QuickCopyApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
