# -*- coding: utf-8 -*-
"""
快捷工作台 · Qt(PySide6) 版
圆角悬浮 + 阴影 + 卡片主题；分组；平滑滚动；邮箱/短信验证码
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
from contextlib import contextmanager
from email.header import decode_header
from email.utils import parsedate_to_datetime

import urllib.request
import zipfile
from PySide6.QtCore import (Qt, QObject, Signal, QPoint, QSize, QEvent,
                            QMimeData)
from PySide6.QtGui import (QFont, QIcon, QColor, QCursor, QDrag)
from PySide6.QtWidgets import (
    QApplication, QWidget, QFrame, QLabel, QPushButton, QVBoxLayout,
    QHBoxLayout, QComboBox, QScrollArea, QSizePolicy, QDialog, QLineEdit,
    QCheckBox, QGridLayout, QGraphicsDropShadowEffect, QInputDialog,
    QFileDialog, QMessageBox)

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(APP_DIR, "quick_copy_data.json")
CONFIG_FILE = os.path.join(APP_DIR, "verify_config.json")

# ---------- 配色 ----------
ACCENT = "#4263eb"
ACCENT_D = "#3651c9"
BG = "#eef1f7"
CARD = "#ffffff"
BORDER = "#e3e7ef"
TEXT = "#232a3a"
MUTED = "#8b94a7"

WIDTH = 320

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

QSS = f"""
QWidget {{ font-family: "Microsoft YaHei UI"; color: {TEXT}; font-size: 9pt; }}
#root {{ background: {BG}; border-radius: 12px; }}
#titlebar {{ background: {ACCENT}; border-top-left-radius:12px;
             border-top-right-radius:12px; }}
#title {{ color:#fff; font-size:10pt; font-weight:700; }}
#tbtn {{ background:transparent; color:#fff; border:none;
         border-radius:7px; padding:4px 8px; font-size:9pt; }}
#tbtn:hover {{ background: rgba(255,255,255,0.20); }}
#tbtn:pressed {{ background: rgba(255,255,255,0.32); }}
.card, #card {{ background:{CARD}; border:1px solid {BORDER}; border-radius:10px; }}
QComboBox {{ background:{CARD}; border:1px solid {BORDER}; border-radius:8px;
             padding:4px 8px; min-height:18px; }}
QComboBox:hover {{ border-color:#b9c9ff; }}
QComboBox::drop-down {{ border:none; width:18px; }}
QComboBox QAbstractItemView {{ background:#fff; border:1px solid {BORDER};
    selection-background-color:#eef3ff; selection-color:{ACCENT};
    outline:none; border-radius:6px; padding:4px; }}
#gbtn {{ background:transparent; border:none; color:{ACCENT};
         font-weight:700; padding:3px 5px; border-radius:6px; }}
#gbtn:hover {{ background:#e6edff; }}
.mini, #mini {{ background:{ACCENT}; color:#fff; border:none; border-radius:7px;
         padding:5px 10px; font-size:8.5pt; }}
.mini:hover, #mini:hover {{ background:{ACCENT_D}; }}
.mini:pressed, #mini:pressed {{ background:#2c45ad; }}
.mini:disabled, #mini:disabled {{ background:#c8cedb; color:#e9ecf3; }}
#code {{ background:#f6f8fe; border:1px solid #dfe6f5; border-radius:8px;
         font-family:Consolas; font-size:18pt; font-weight:700;
         color:{ACCENT}; padding:4px; }}
.row, #row {{ background:#fff; border:1px solid #e6eaf2; border-radius:8px; }}
.row:hover, #row:hover {{ background:#f1f5ff; border-color:#c2d2ff; }}
.row QLabel, #row QLabel {{ background:transparent; }}
#rbtn {{ background:transparent; border:none; color:{MUTED};
         padding:2px 4px; border-radius:5px; }}
#rbtn:hover {{ background:#e7edff; color:{ACCENT}; }}
QScrollArea {{ border:none; background:transparent; }}
#body,#rowshost {{ background:transparent; }}
QScrollBar:vertical {{ background:transparent; width:8px; margin:2px; }}
QScrollBar::handle:vertical {{ background:#cfd6e4; border-radius:4px;
                               min-height:36px; }}
QScrollBar::handle:vertical:hover {{ background:#aab4c8; }}
QScrollBar::add-line,QScrollBar::sub-line {{ height:0; }}
QScrollBar::add-sub,QScrollBar::sub-page {{ background:transparent; }}
#footer {{ color:{MUTED}; font-size:8pt; }}
QLineEdit,QSpin {{ background:#fff; border:1px solid {BORDER};
    border-radius:7px; padding:5px 8px; }}
QLineEdit:focus {{ border-color:{ACCENT}; }}
QCheckBox {{ spacing:6px; }}
QDialog,QMessageBox {{ background:{BG}; }}
#dlgbtn {{ background:{ACCENT}; color:#fff; border:none; border-radius:7px;
           padding:6px 14px; }}
#dlgbtn:hover {{ background:{ACCENT_D}; }}
#dlgbtn2 {{ background:#e7eaf1; color:#555; border:none; border-radius:7px;
            padding:6px 14px; }}
#dlgbtn2:hover {{ background:#d9deea; }}
#adbbtn {{ background:{ACCENT}; color:#fff; border:none; border-radius:7px;
           padding:6px 2px; font-size:12px; min-width:54px; }}
#adbbtn:hover {{ background:{ACCENT_D}; }}
#swatch {{ border:1px solid #d5dbe6; border-radius:6px; max-width:26px;
           min-height:22px; }}
QToolTip {{ background:#2b3245; color:#ffffff; border:none;
            border-radius:5px; padding:4px 8px; }}
"""


# ================= 后端：DPAPI =================
class _Blob(ctypes.Structure):
    _fields_ = [("cbData", ctypes.wintypes.DWORD),
                ("pbData", ctypes.POINTER(ctypes.c_char))]


def dpapi_protect(text):
    raw = text.encode("utf-8")
    inb = _Blob(len(raw), ctypes.cast(ctypes.c_char_p(raw),
                                      ctypes.POINTER(ctypes.c_char)))
    out = _Blob()
    if not ctypes.windll.crypt32.CryptProtectData(
            ctypes.byref(inb), None, None, None, None, 0, ctypes.byref(out)):
        raise ctypes.WinError()
    b = ctypes.string_at(out.pbData, out.cbData)
    ctypes.windll.kernel32.LocalFree(out.pbData)
    return base64.b64encode(b).decode()


def dpapi_unprotect(b64):
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


# ================= 后端：邮件 =================
def _decoded_str(raw):
    if raw is None:
        return ""
    out = []
    for t, enc in decode_header(raw):
        if isinstance(t, bytes):
            out.append(t.decode(enc or "utf-8", errors="replace"))
        else:
            out.append(t)
    return "".join(out)


def _msg_text(msg):
    chunks = []
    for part in (msg.walk() if msg.is_multipart() else [msg]):
        if msg.is_multipart() and part.get_content_type() not in (
                "text/plain", "text/html"):
            continue
        p = part.get_payload(decode=True)
        if p:
            chunks.append(p.decode(part.get_content_charset() or "utf-8",
                                   errors="replace"))
    return re.sub(r"<[^>]+>", " ", "\n".join(chunks))


def extract_code(text):
    m = CODE_NEAR_RE.search(text)
    if m:
        return m.group(1)
    c = NUM_RE.findall(text)
    return c[0] if c else None


@contextmanager
def ipv4_only():
    orig = socket.getaddrinfo

    def f(host, port=0, family=0, type_=0, proto=0, flags=0):
        return [i for i in orig(host, port, family, type_, proto, flags)
                if i[0] == socket.AF_INET]
    socket.getaddrinfo = f
    try:
        yield
    finally:
        socket.getaddrinfo = orig


def _netease(M, cfg):
    M.login(cfg["addr"], dpapi_unprotect(cfg["secret"]))
    if any(k in cfg["host"] for k in ("163.com", "126.com", "netease")):
        # imaplib 的命令状态表默认不含扩展命令 ID，需先登记否则 KeyError
        imaplib.Commands.setdefault("ID", ("AUTH", "SELECTED", "LOGOUT"))
        payload = '("name" "QuickDesk" "version" "3" "vendor" "local")'
        t, d = M._simple_command("ID", payload)
        M._untagged_response(t, d, "ID")


def _inbox(M):
    t, d = M.select("INBOX", readonly=True)
    if t != "OK":
        raise RuntimeError(f"打开收件箱失败 {d}")


def imap_baseline(cfg):
    with ipv4_only():
        M = imaplib.IMAP4_SSL(cfg["host"], 993, timeout=12)
    _netease(M, cfg); _inbox(M)
    t, d = M.uid("search", "ALL")
    if t != "OK":
        raise RuntimeError(f"SEARCH 失败 {d}")
    u = [int(x) for x in d[0].split() if x.isdigit()]
    M.logout()
    return max(u) if u else 0


def imap_fetch(cfg):
    base = int(cfg.get("base_uid", 0))
    with ipv4_only():
        M = imaplib.IMAP4_SSL(cfg["host"], 993, timeout=12)
    _netease(M, cfg); _inbox(M)
    t, d = M.uid("search", "UID", f"{base+1}:*")
    if t != "OK":
        M.logout(); raise RuntimeError(f"SEARCH 失败 {d}")
    uids = [x for x in d[0].split() if x.isdigit()]
    for uid in reversed(uids):
        if int(uid) <= base:
            continue
        t, md = M.uid("fetch", uid, "(RFC822)")
        if not md or not md[0]:
            continue
        msg = email.message_from_bytes(md[0][1])
        code = extract_code(_decoded_str(msg.get("Subject")) + "\n"
                            + _msg_text(msg))
        if code:
            try:
                ts = parsedate_to_datetime(msg.get("Date")).timestamp()
            except Exception:
                ts = time.time()
            M.logout()
            return int(uid), code, ts
    M.logout()
    return None


# ================= 后端：ADB =================
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
    for line in out.stdout.decode("utf-8", "replace").splitlines()[1:]:
        p = line.strip().split("\t")
        if len(p) == 2 and p[1] == "device":
            devs.append(p[0])
    return adb, devs


def sms_fetch(cfg):
    adb = find_adb()
    cmd = [adb]
    if cfg.get("serial"):
        cmd += ["-s", cfg["serial"]]
    cmd += ["shell", "content", "query", "--uri", "content://sms/inbox",
            "--projection", "date,body", "--sort", "date DESC LIMIT 15"]
    out = subprocess.run(cmd, capture_output=True, timeout=12)
    base = int(cfg.get("base_date", 0))
    for line in out.stdout.decode("utf-8", "replace").splitlines():
        m = re.match(r"Row:\s*\d+\s*date=(\d+)\s*body=(.*)", line, re.S)
        if not m:
            continue
        dms, body = int(m.group(1)), m.group(2).rstrip(",")
        if dms <= base:
            continue
        code = extract_code(body)
        if code:
            return dms, code
    return None


def install_adb(cb):
    urls = [
        "https://dl.google.com/android/repository/platform-tools-latest-windows.zip",
        "https://mirrors.bfsu.edu.cn/android/repository/platform-tools-latest-windows.zip",
        "https://mirrors.tuna.tsinghua.edu.cn/android/repository/platform-tools-latest-windows.zip"]
    zp = os.path.join(APP_DIR, "platform-tools.zip")
    last = None
    for u in urls:
        try:
            cb(f"下载中 {u.split('/')[2]}…")
            req = urllib.request.Request(u,
                                         headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=60) as r, \
                    open(zp, "wb") as f:
                f.write(r.read())
            break
        except Exception as e:
            last = e
    else:
        raise RuntimeError(f"全部源失败 {last}")
    cb("解压中…")
    with zipfile.ZipFile(zp) as z:
        z.extractall(APP_DIR)
    dst = os.path.join(APP_DIR, "adb")
    if os.path.exists(dst):
        shutil.rmtree(dst)
    os.rename(os.path.join(APP_DIR, "platform-tools"), dst)
    os.remove(zp)
    cb("adb 安装完成")


# ================= 字段编辑对话框 =================
COLOR_CHOICES = ["#333333", "#e53935", ACCENT, "#2e9b46",
                 "#f08c1a", "#8e44ad"]


class FieldDialog(QDialog):
    def __init__(self, parent, item=None):
        super().__init__(parent)
        self.setWindowTitle("编辑条目")
        self.setFixedWidth(290)
        self.picked = item.get("color", "#666666") if item else "#666666"
        lay = QVBoxLayout(self); lay.setSpacing(8); lay.setContentsMargins(14,14,14,14)

        self.e_name = QLineEdit(item["name"] if item else "")
        self.e_name.setPlaceholderText("字段名（组内唯一）")
        self.e_val = QLineEdit(item["value"] if item else "")
        self.e_val.setPlaceholderText("内容")
        lay.addWidget(self.e_name); lay.addWidget(self.e_val)

        row = QHBoxLayout()
        self.cb_bold = QCheckBox("加粗 key")
        self.cb_bold.setChecked(bool(item and item.get("bold")))
        row.addWidget(self.cb_bold); row.addStretch()
        self.sw = {}
        for c in COLOR_CHOICES:
            b = QPushButton(); b.setObjectName("swatch")
            b.setStyleSheet(f"background:{c};")
            b.setCheckable(True); b.setToolTip("key 显示颜色")
            b.clicked.connect(lambda _, col=c: self.pick(col))
            row.addWidget(b); self.sw[c] = b
        lay.addLayout(row)
        self.pick(self.picked)

        bs = QHBoxLayout(); bs.addStretch()
        no = QPushButton("取消"); no.setObjectName("dlgbtn2")
        no.setToolTip("放弃本次修改")
        no.clicked.connect(self.reject)
        ok = QPushButton("确定"); ok.setObjectName("dlgbtn")
        ok.setToolTip("保存此条目")
        ok.clicked.connect(self.accept)
        bs.addWidget(no); bs.addWidget(ok)
        lay.addLayout(bs)

    def pick(self, c):
        self.picked = c
        for col, b in self.sw.items():
            b.setChecked(col == c)
            b.setText("✓" if col == c else "")
            b.setStyleSheet(
                f"background:{col}; color:#fff;"
                f"border:2px solid {'#222' if col==c else '#d5dbe6'};")

    def result_item(self):
        return {"name": self.e_name.text().strip() or "未命名",
                "value": self.e_val.text().strip(),
                "bold": self.cb_bold.isChecked(), "color": self.picked}


# ================= 绑定对话框 =================
class BindDialog(QDialog):
    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.setWindowTitle("绑定邮箱 / 手机")
        self.setFixedWidth(440)
        lay = QVBoxLayout(self); lay.setSpacing(7); lay.setContentsMargins(16,14,16,14)

        lay.addWidget(QLabel("① 邮箱（自动获取邮件验证码）"))
        self.e_addr = QLineEdit(); self.e_addr.setPlaceholderText("邮箱地址")
        self.e_sec = QLineEdit(); self.e_sec.setPlaceholderText("IMAP 授权码（非登录密码）")
        self.e_host = QLineEdit(); self.e_host.setPlaceholderText("IMAP 服务器")
        old = app.verify_cfg.get("email")
        if old:
            self.e_addr.setText(old.get("addr",""))
            self.e_host.setText(old.get("host",""))
        self.e_addr.textChanged.connect(self.auto_host)
        lay.addWidget(self.e_addr); lay.addWidget(self.e_sec); lay.addWidget(self.e_host)
        b1 = QPushButton("保存并绑定邮箱"); b1.setObjectName("dlgbtn")
        b1.setToolTip("登录邮箱并建立收件基线")
        b1.clicked.connect(self.save_email); lay.addWidget(b1)
        self.mail_st = QLabel(""); self.mail_st.setStyleSheet(f"color:{ACCENT};")
        lay.addWidget(self.mail_st)

        lay.addWidget(QLabel("② 安卓手机（短信验证码 / USB 或无线）"))
        tip2 = QLabel("无线步骤：手机点“使用配对码配对设备”→ 填“配对”三项后点配对；成功后用主页面端口填“连接”再点连接")
        tip2.setStyleSheet(f"color:{MUTED}; font-size:11px;"); tip2.setWordWrap(True)
        lay.addWidget(tip2)
        g = QGridLayout(); g.setSpacing(6)
        self.e_pair_ip = QLineEdit(); self.e_pair_ip.setPlaceholderText("配对 IP，如 192.168.5.22")
        self.e_pair_port = QLineEdit(); self.e_pair_port.setPlaceholderText("配对端口")
        self.e_pcode = QLineEdit(); self.e_pcode.setPlaceholderText("6位配对码")
        self.e_conn_ip = QLineEdit(); self.e_conn_ip.setPlaceholderText("连接 IP，一般同左")
        self.e_conn_port = QLineEdit(); self.e_conn_port.setPlaceholderText("连接端口（主页面）")
        g.addWidget(QLabel("配对"),0,0)
        g.addWidget(self.e_pair_ip,0,1); g.addWidget(self.e_pair_port,0,2)
        g.addWidget(self.e_pcode,1,1,1,2)
        g.addWidget(QLabel("连接"),2,0)
        g.addWidget(self.e_conn_ip,2,1); g.addWidget(self.e_conn_port,2,2)
        self.e_pair_ip.textChanged.connect(
            lambda t: self.e_conn_ip.setText(t) if not self.e_conn_ip.text() else None)
        self.sms_st = QLabel(""); self.sms_st.setStyleSheet(f"color:{MUTED};")
        lay.addLayout(g)
        r = QHBoxLayout(); r.setSpacing(5)
        bspecs = (("检测", self.check, "检测 adb 与已连接设备"),
                  ("配对", self.pair, "无线配对（填配对页地址端口+配对码）"),
                  ("连接", self.connect, "无线连接设备（配对成功后）"),
                  ("下载adb", self.dl_adb, "自动下载安装 adb"),
                  ("绑定设备", self.bind, "绑定当前设备用于读短信"))
        for t, fn, tip in bspecs:
            b = QPushButton(t); b.setObjectName("adbbtn"); b.setToolTip(tip)
            b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(fn); r.addWidget(b)
        lay.addLayout(r)
        lay.addWidget(self.sms_st)

        close = QPushButton("关闭"); close.setObjectName("dlgbtn2")
        close.setToolTip("关闭绑定窗口")
        close.clicked.connect(self.accept)
        lay.addWidget(close)
        self.check()

    def auto_host(self, t):
        if "@" in t:
            h = EMAIL_HOSTS.get(t.split("@",1)[1].lower())
            if h and not self.e_host.text():
                self.e_host.setText(h)

    def save_email(self):
        a, s, h = self.e_addr.text().strip(), self.e_sec.text().strip(), \
                  self.e_host.text().strip()
        if not (a and s and h):
            self.mail_st.setText("三项都要填"); self.mail_st.setStyleSheet("color:#e53935;")
            return
        cfg = {"addr": a, "host": h, "secret": dpapi_protect(s),
               "base_uid": 0}
        self.mail_st.setText("登录建立基线…")

        def work():
            try:
                cfg["base_uid"] = imap_baseline(cfg)
                self.app.verify_cfg["email"] = cfg
                self.app.save_verify_cfg()
                self.mail_st.setText(f"绑定成功 基线UID={cfg['base_uid']}")
                self.mail_st.setStyleSheet(f"color:{ACCENT};")
            except Exception as ex:
                self.mail_st.setText(f"失败：{str(ex)[:50]}")
                self.mail_st.setStyleSheet("color:#e53935;")
        threading.Thread(target=work, daemon=True).start()

    def check(self):
        self.sms_st.setText("检测中…")
        def work():
            _, devs = adb_devices()
            if not find_adb():
                t, c = "未检测到adb，可下载", "#e53935"
            elif not devs:
                t, c = "adb就绪，暂无设备", "#e53935"
            else:
                t, c = "已连接：" + "、".join(devs), ACCENT
            self.sms_st.setText(t); self.sms_st.setStyleSheet(f"color:{c};")
        threading.Thread(target=work, daemon=True).start()

    def _adb(self, args):
        adb = find_adb()
        if not adb:
            self.sms_st.setText("未装adb"); return
        def work():
            try:
                out = subprocess.run([adb]+args, capture_output=True, timeout=30)
                t = (out.stdout.decode("utf-8","replace") +
                     out.stderr.decode("utf-8","replace")).strip()
                self.sms_st.setText(t[:90] or "完成")
            except Exception as e:
                self.sms_st.setText(str(e)[:60])
        threading.Thread(target=work, daemon=True).start()

    @staticmethod
    def _valid_port(p):
        return p.isdigit() and 1 <= int(p) <= 65535

    def _warn(self, m):
        self.sms_st.setText(m); self.sms_st.setStyleSheet("color:#e53935;")

    def pair(self):
        ip = self.e_pair_ip.text().strip()
        port = self.e_pair_port.text().strip()
        code = self.e_pcode.text().strip()
        if not (ip and port and code):
            self._warn("配对需填：配对 IP、配对端口、6位配对码"); return
        if not self._valid_port(port):
            self._warn(f"配对端口无效：{port}（1-65535，一般5位）"); return
        self._adb(["pair", f"{ip}:{port}", code])

    def connect(self):
        ip = self.e_conn_ip.text().strip() or self.e_pair_ip.text().strip()
        port = self.e_conn_port.text().strip()
        if not (ip and port):
            self._warn("连接需填：连接 IP 和主页面上的连接端口"); return
        if not self._valid_port(port):
            self._warn(f"连接端口无效：{port}（1-65535，一般5位）"); return
        self._adb(["connect", f"{ip}:{port}"])

    def dl_adb(self):
        def work():
            try:
                install_adb(lambda m: self.sms_st.setText(m))
            except Exception as e:
                self.sms_st.setText("失败："+str(e)[:45])
        threading.Thread(target=work, daemon=True).start()

    def bind(self):
        _, devs = adb_devices()
        if not devs:
            self.sms_st.setText("没有已连接设备"); return
        self.app.verify_cfg["sms"] = {"serial": devs[0],
                                      "base_date": int(time.time()*1000)}
        self.app.save_verify_cfg()
        self.sms_st.setText(f"已绑定 {devs[0]}")


# ================= 拖放宿主 =================
class DropHost(QWidget):
    requestMove = Signal(int, int)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)

    def dragEnterEvent(self, e):
        if e.mimeData().hasFormat("application/x-rowidx"):
            e.acceptProposedAction()

    def dragMoveEvent(self, e):
        if e.mimeData().hasFormat("application/x-rowidx"):
            e.acceptProposedAction()

    def dropEvent(self, e):
        if not e.mimeData().hasFormat("application/x-rowidx"):
            return
        src = int(bytes(e.mimeData().data("application/x-rowidx")).decode())
        cards = self.findChildren(RowCard)
        target = len(cards)
        y = e.position().y()
        for i, c in enumerate(cards):
            cy = c.y() + c.height()/2
            if y < cy:
                target = i; break
        e.acceptProposedAction()
        self.requestMove.emit(src, target)


# ================= 字段行卡片 =================
class RowCard(QFrame):
    def __init__(self, main, idx, item):
        super().__init__()
        self.setObjectName("row")
        self.main, self.idx = main, idx
        h = QHBoxLayout(self); h.setContentsMargins(8,3,6,3); h.setSpacing(6)

        key_font = QFont("Microsoft YaHei UI", 9)
        key_font.setBold(bool(item.get("bold")))
        self.k = QLabel(item["name"]); self.k.setFont(key_font)
        self.k.setStyleSheet(f"color:{item.get('color','#666')};")
        self.k.setFixedWidth(78); self.k.setAlignment(Qt.AlignRight
                                                      | Qt.AlignVCenter)
        self.v = QLabel(); self.v.setMinimumWidth(90)
        self._set_value_text(item)
        b_edit = QPushButton("✎"); b_edit.setObjectName("rbtn")
        b_del = QPushButton("✕"); b_del.setObjectName("rbtn")
        b_edit.setToolTip("修改此条目")
        b_del.setToolTip("删除此条目（有确认）")
        b_edit.clicked.connect(lambda: main.click_edit(idx))
        b_del.clicked.connect(lambda: main.click_delete(idx, b_del))
        h.addWidget(self.k); h.addWidget(self.v, 1)
        h.addWidget(b_edit); h.addWidget(b_del)
        self.setToolTip("点击复制 · 按住可拖动排序")
        self._press = None; self._moved = False
        self.setCursor(Qt.PointingHandCursor)

    def _set_value_text(self, item):
        if self.main.masked:
            self.v.setText("••••••" if item["value"] else "（空）")
        else:
            val = item["value"] or "（空，点✎填写）"
            self.v.setText(val[:24] + ("…" if len(val) > 24 else ""))

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._press = e.globalPosition().toPoint()
            self._moved = False

    def mouseMoveEvent(self, e):
        if self._press is None:
            return
        if (e.globalPosition().toPoint()-self._press).manhattanLength() < 8:
            return
        self._moved = True
        drag = QDrag(self)
        mime = QMimeData()
        mime.setData("application/x-rowidx", str(self.idx).encode())
        drag.setMimeData(mime)
        drag.exec(Qt.MoveAction)

    def mouseReleaseEvent(self, e):
        if e.button() == Qt.LeftButton and not self._moved:
            self.main.do_copy(self.idx)
        self._press = None


# ================= 主窗口 =================
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.groups = []
        self.active_group = 0
        self.pos_xy = None
        self.collapsed = False
        self.pin_on = True
        self.masked = False
        self.verify_cfg = {"show_panel": True, "email": None, "sms": None}
        self._drag = None
        self._watch = None; self._remain = 0
        self._current_code = ""
        self.rows = []

        self.load_data(); self.load_verify_cfg()

        self.setWindowFlags(Qt.FramelessWindowHint |
                            (Qt.WindowStaysOnTopHint if self.pin_on else 0))
        self.setFixedWidth(WIDTH)

        outer = QVBoxLayout(self); outer.setContentsMargins(0,0,0,0)
        self.root = QFrame(); self.root.setObjectName("root")
        outer.addWidget(self.root)

        rl = QVBoxLayout(self.root); rl.setContentsMargins(0,0,0,8)
        rl.setSpacing(4)
        self.build_titlebar(rl)
        self.body = QFrame(); self.body.setObjectName("body")
        bl = QVBoxLayout(self.body); bl.setContentsMargins(10,4,10,0)
        bl.setSpacing(6)
        self.build_groupbar(bl)
        self.build_vpanel(bl)
        self.build_rows_area(bl)
        self.footer = QLabel("点条目复制 · 数据本地保存")
        self.footer.setObjectName("footer")
        bl.addWidget(self.footer)
        rl.addWidget(self.body)

        self.refresh_groups(); self.refresh_rows(); self.adjust_height()
        self.update_btns_state()
        if not self.verify_cfg.get("show_panel", True):
            self.vpanel.hide()
        if self.collapsed:
            self.body.hide()
        x, y = self.pos_xy or (90, 120)
        self.move(x, y)

        # Win11 圆角 + 原生阴影（给无边框窗口补回 CAPTION/THICKFRAME 样式）
        try:
            hwnd = int(self.winId())
            GWL_STYLE = -16
            WS_CAPTION, WS_THICKFRAME = 0x00C00000, 0x00040000
            cur = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_STYLE)
            ctypes.windll.user32.SetWindowLongW(
                hwnd, GWL_STYLE, cur | WS_CAPTION | WS_THICKFRAME)
            DWM_WCP = 33; DWMWCP_ROUND = 2
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, DWM_WCP,
                ctypes.byref(ctypes.c_int(DWMWCP_ROUND)),
                ctypes.sizeof(ctypes.c_int))
        except Exception:
            pass

        def warm():
            a = find_adb()
            if a: subprocess.run([a,"start-server"], capture_output=True)
        threading.Thread(target=warm, daemon=True).start()

    # ---------- 标题栏 ----------
    def build_titlebar(self, parent_layout):
        tb = QFrame(); tb.setObjectName("titlebar"); tb.setFixedHeight(40)
        h = QHBoxLayout(tb); h.setContentsMargins(12,4,8,4)
        title = QLabel("快捷工作台"); title.setObjectName("title")
        title.setAttribute(Qt.WA_TransparentForMouseEvents)
        h.addWidget(title); h.addStretch()
        self.t_add = self.tb_btn("＋", self.click_add)
        self.t_mask = self.tb_btn("隐", self.toggle_mask)
        self.t_code = self.tb_btn("码", self.toggle_vpanel)
        self.t_pin = self.tb_btn("📌", self.toggle_pin)
        self.t_col = self.tb_btn("▼", self.toggle_collapse)
        self.t_close = self.tb_btn("✕", self.close_win)
        for b in (self.t_close,self.t_col,self.t_pin,self.t_code,
                  self.t_mask,self.t_add):
            h.addWidget(b)
        self.t_add.setToolTip("添加一条新字段")
        self.t_mask.setToolTip("隐藏 / 显示所有条目内容")
        self.t_code.setToolTip("显示 / 隐藏验证码助手")
        self.t_pin.setToolTip("窗口始终置顶开关")
        self.t_col.setToolTip("折叠 / 展开窗口")
        self.t_close.setToolTip("关闭程序")
        self.t_mask.setText("显" if self.masked else "隐")
        parent_layout.addWidget(tb)
        self.titlebar = tb
        tb.installEventFilter(self)

    def eventFilter(self, obj, ev):
        if obj is self.titlebar:
            if ev.type() == QEvent.MouseButtonPress and ev.button()==Qt.LeftButton:
                self._drag = ev.globalPosition().toPoint()-self.pos()
            elif ev.type() == QEvent.MouseMove and self._drag is not None:
                self.move(ev.globalPosition().toPoint()-self._drag)
            elif ev.type() == QEvent.MouseButtonRelease:
                self._drag = None
                self.pos_xy = [self.x(), self.y()]; self.save_data()
        return False

    def tb_btn(self, text, fn):
        b = QPushButton(text); b.setObjectName("tbtn")
        b.setCursor(Qt.PointingHandCursor); b.clicked.connect(fn)
        return b

    # ---------- 分组栏 ----------
    def build_groupbar(self, layout):
        g = QHBoxLayout(); g.setSpacing(2)
        self.cb = QComboBox()
        self.cb.setMinimumHeight(28)
        self.cb.currentIndexChanged.connect(self.on_group)
        self.cb.setToolTip("切换分组")
        g.addWidget(self.cb, 1)
        gspecs = (("＋", self.add_group, "新建一个分组"),
                  ("✎", self.rename_group, "给当前分组改名"),
                  ("✕", self.delete_group, "删除当前分组（有确认）"),
                  ("⇓", self.import_data, "从 JSON 文件导入数据"),
                  ("⇑", self.export_data, "导出 JSON 备份 / 换机用"))
        for t, fn, tip in gspecs:
            b = QPushButton(t); b.setObjectName("gbtn")
            b.setCursor(Qt.PointingHandCursor); b.clicked.connect(fn)
            b.setToolTip(tip); g.addWidget(b)
        layout.addLayout(g)

    def refresh_groups(self):
        self.cb.blockSignals(True)
        self.cb.clear()
        for g in self.groups:
            self.cb.addItem(g["name"])
        self.cb.setCurrentIndex(self.active_group)
        self.cb.blockSignals(False)

    def on_group(self, i):
        self.active_group = i
        self.refresh_rows(); self.save_data()

    @property
    def fields(self):
        return self.groups[self.active_group]["fields"]

    def add_group(self):
        nm, ok = QInputDialog.getText(self, "新建组", "组名：")
        nm = nm.strip()
        if ok and nm:
            if any(g["name"] == nm for g in self.groups):
                self.flash("组名已存在", True); return
            self.groups.append({"name": nm, "fields": []})
            self.active_group = len(self.groups)-1
            self.refresh_groups(); self.refresh_rows(); self.save_data()

    def rename_group(self):
        cur = self.groups[self.active_group]
        nm, ok = QInputDialog.getText(self, "改名", "新组名：", text=cur["name"])
        nm = nm.strip()
        if ok and nm:
            self.groups[self.active_group]["name"] = nm
            self.refresh_groups(); self.save_data()

    def delete_group(self):
        if len(self.groups) <= 1:
            self.flash("至少保留一个组", True); return
        cur = self.groups[self.active_group]
        if self.confirm_delete(
                f"确定删除整组「{cur['name']}」吗？",
                f"组内 {len(cur['fields'])} 个条目会一起删除，且无法恢复。"):
            del self.groups[self.active_group]
            self.active_group = max(0, self.active_group-1)
            self.refresh_groups(); self.refresh_rows(); self.save_data()

    def confirm_delete(self, title, detail=""):
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Warning)
        box.setWindowTitle("确认删除")
        box.setText(title)
        if detail:
            box.setInformativeText(detail)
        box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        box.setDefaultButton(QMessageBox.No)
        box.button(QMessageBox.Yes).setText("删除")
        box.button(QMessageBox.No).setText("取消")
        return box.exec() == QMessageBox.Yes

    # ---------- 验证码面板 ----------
    def build_vpanel(self, layout):
        card = QFrame(); card.setObjectName("card")
        v = QVBoxLayout(card); v.setContentsMargins(10,8,10,10); v.setSpacing(6)
        v.addWidget(QLabel("验证码助手"))
        self.code_lbl = QLabel("———"); self.code_lbl.setObjectName("code")
        self.code_lbl.setAlignment(Qt.AlignCenter)
        self.code_lbl.setCursor(Qt.PointingHandCursor)
        self.code_lbl.setToolTip("点击复制验证码")
        self.code_lbl.mousePressEvent = lambda e: self.copy_code()
        v.addWidget(self.code_lbl)
        r = QHBoxLayout()
        self.b_email = QPushButton("邮箱")
        self.b_sms = QPushButton("短信")
        b_bind = QPushButton("绑定")
        b_stop = QPushButton("停止")
        specs = ((self.b_email, self.fetch_email, "从已绑定邮箱获取新验证码"),
                 (self.b_sms, self.fetch_sms, "从已绑定手机短信获取验证码"),
                 (b_bind, self.open_bind, "绑定邮箱 / 手机"),
                 (b_stop, self.stop_watch, "停止等待验证码"))
        for b, fn, tip in specs:
            b.setObjectName("mini"); b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(fn); b.setToolTip(tip); r.addWidget(b)
        v.addLayout(r)
        self.code_st = QLabel("未获取"); self.code_st.setStyleSheet(f"color:{MUTED};")
        v.addWidget(self.code_st)
        self.vpanel = card
        layout.addWidget(card)

    def toggle_vpanel(self):
        if self.vpanel.isVisible():
            self.vpanel.hide(); self.verify_cfg["show_panel"] = False
            self.stop_watch()
        else:
            self.vpanel.show(); self.verify_cfg["show_panel"] = True
        self.save_verify_cfg(); self.adjust_height()

    def code_status(self, t, c=MUTED):
        self.code_st.setText(t); self.code_st.setStyleSheet(f"color:{c};")

    def copy_code(self):
        if not self._current_code:
            self.code_status("还没有验证码", "#e53935"); return
        QApplication.clipboard().setText(self._current_code)
        self.code_status(f"已复制 {self._current_code}", ACCENT)

    def fetch_email(self):
        cfg = self.verify_cfg.get("email")
        if not cfg:
            self.code_status("请先绑定邮箱", "#e53935"); self.open_bind(); return
        self.code_status("查询邮箱中…")
        self.start_watch("email")
        threading.Thread(target=self.email_worker, args=(cfg,),
                         daemon=True).start()

    def email_worker(self, cfg):
        try:
            r = imap_fetch(cfg)
            if r:
                uid, code, ts = r
                self._found(code, "email", ts, uid)
        except Exception as e:
            self._fail("邮箱", str(e))

    def fetch_sms(self):
        cfg = self.verify_cfg.get("sms")
        if not cfg:
            self.code_status("请先绑定手机", "#e53935"); self.open_bind(); return
        _, devs = adb_devices()
        if not devs:
            self.code_status("手机未连接", "#e53935"); return
        self.code_status("读取短信中…")
        self.start_watch("sms")
        threading.Thread(target=self.sms_worker, args=(cfg,),
                         daemon=True).start()

    def sms_worker(self, cfg):
        try:
            r = sms_fetch(cfg)
            if r:
                dms, code = r
                self._found(code, "sms", dms/1000, dms)
        except Exception as e:
            self._fail("短信", str(e))

    def _found(self, code, src, ts, advance):
        self._current_code = code
        self.code_lbl.setText(code)
        self.stop_watch()
        QApplication.clipboard().setText(code)
        if src == "email":
            self.verify_cfg["email"]["base_uid"] = advance
        else:
            self.verify_cfg["sms"]["base_date"] = advance
        self.save_verify_cfg()
        ago = max(0, int(time.time()-ts))
        tag = "邮箱" if src=="email" else "短信"
        self.code_status(f"{tag}验证码已获取并自动复制（{ago}秒前）", ACCENT)

    def start_watch(self, src):
        self.stop_watch()
        self._watch_src = src
        self._remain = 60      # 总等待 60 秒
        self._poll_tick = 0
        self.code_status("等待新验证码…（60s）")
        self._watch = threading.Timer(1.0, self._watch_tick)
        self._watch.start()

    def _watch_tick(self):
        self._remain -= 1
        self._poll_tick += 1
        if self._remain <= 0:
            self.stop_watch()
            self.code_status("等待超时：未收到新验证码", "#e53935")
            return
        self.code_status(f"等待新验证码…（{self._remain}s）")
        if self._poll_tick % 3 == 0:   # 后台每 3 秒实际查询一次
            self._retry()
        self._watch = threading.Timer(1.0, self._watch_tick)
        self._watch.start()

    def _retry(self):
        if self._watch_src == "email":
            cfg = self.verify_cfg.get("email")
            if cfg: threading.Thread(target=self.email_worker,args=(cfg,),daemon=True).start()
        else:
            cfg = self.verify_cfg.get("sms")
            if cfg: threading.Thread(target=self.sms_worker,args=(cfg,),daemon=True).start()

    def _fail(self, tag, msg):
        self.stop_watch(); self.code_status(f"{tag}失败：{msg[:45]}", "#e53935")

    def stop_watch(self):
        if self._watch: self._watch.cancel(); self._watch = None
        self._remain = 0
        self.code_status("已停止等待")

    def update_btns_state(self):
        self.b_email.setEnabled(bool(self.verify_cfg.get("email")))
        self.b_sms.setEnabled(bool(self.verify_cfg.get("sms")))

    def open_bind(self):
        BindDialog(self).exec()
        self.update_btns_state()

    # ---------- 行区域 ----------
    def build_rows_area(self, layout):
        self.scroll = QScrollArea(); self.scroll.setWidgetResizable(True)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.rows_host = DropHost(); self.rows_host.setObjectName("rowshost")
        self.rows_host.requestMove.connect(self.move_row)
        self.rows_lay = QVBoxLayout(self.rows_host)
        self.rows_lay.setContentsMargins(2,2,6,2); self.rows_lay.setSpacing(4)
        self.rows_lay.addStretch()
        self.scroll.setWidget(self.rows_host)
        layout.addWidget(self.scroll)

    def refresh_rows(self):
        while self.rows_lay.count() > 1:
            it = self.rows_lay.takeAt(0)
            w = it.widget()
            if w: w.deleteLater()
        self.rows = []
        if not self.fields:
            empty = QLabel("本组还没有条目\n点标题栏 ＋ 添加")
            empty.setAlignment(Qt.AlignCenter)
            empty.setStyleSheet(f"color:{MUTED}; padding:16px;")
            self.rows_lay.insertWidget(0, empty)
        else:
            for i, item in enumerate(self.fields):
                rc = RowCard(self, i, item)
                self.rows_lay.insertWidget(i, rc); self.rows.append(rc)
        self.adjust_height()

    def move_row(self, src, target):
        f = self.fields
        if not (0 <= src < len(f)) or target == src:
            return
        item = f.pop(src)
        if target > src:
            target -= 1
        f.insert(target, item)
        self.refresh_rows(); self.save_data()

    # ---------- 复制/增改删 ----------
    def do_copy(self, idx):
        t = self.fields[idx]["value"]
        if not t:
            self.flash("该条目为空，点 ✎ 填写", True); return
        QApplication.clipboard().setText(t)
        self.flash(f"已复制【{self.fields[idx]['name']}】")

    def click_add(self):
        dlg = FieldDialog(self)
        if dlg.exec():
            item = dlg.result_item()
            if any(f["name"] == item["name"] for f in self.fields):
                self.flash(f"“{item['name']}”已存在，必须唯一", True); return
            self.fields.append(item); self.refresh_rows(); self.save_data()

    def click_edit(self, idx):
        dlg = FieldDialog(self, self.fields[idx])
        if dlg.exec():
            item = dlg.result_item()
            for i, f in enumerate(self.fields):
                if i != idx and f["name"] == item["name"]:
                    self.flash("组内名称重复", True); return
            self.fields[idx] = item; self.refresh_rows(); self.save_data()

    def click_delete(self, idx, btn):
        item = self.fields[idx]
        if self.confirm_delete(f"确定删除条目「{item['name']}」吗？",
                               "删除后无法恢复。"):
            del self.fields[idx]; self.refresh_rows(); self.save_data()

    # ---------- 窗口行为 ----------
    def adjust_height(self):
        QApplication.processEvents()
        n = max(len(self.fields), 1)
        rows_h = min(n*38+8, 330)
        h = 40 + 8
        if not self.collapsed:
            h += rows_h
            if self.verify_cfg.get("show_panel", True):
                h += 132
            h += 34
        self.setFixedHeight(h+24)

    def toggle_collapse(self):
        self.collapsed = not self.collapsed
        self.body.setVisible(not self.collapsed)
        self.t_col.setText("▲" if self.collapsed else "▼")
        self.adjust_height(); self.save_data()

    def toggle_pin(self):
        self.pin_on = not self.pin_on
        flags = Qt.FramelessWindowHint | (Qt.WindowStaysOnTopHint
                                          if self.pin_on else 0)
        self.setWindowFlags(flags)
        self.show()
        self.save_data()

    def toggle_mask(self):
        self.masked = not self.masked
        self.t_mask.setText("显" if self.masked else "隐")
        self.refresh_rows(); self.save_data()

    def flash(self, t, warn=False):
        self.footer.setText(t)
        self.footer.setStyleSheet(f"color:{'#e53935' if warn else ACCENT};")
        def back():
            self.footer.setText("点条目复制 · 数据本地保存")
            self.footer.setStyleSheet(f"color:{MUTED};")
        threading.Timer(1.6, back).start()

    # 拖动
    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag = e.globalPosition().toPoint()-self.pos()

    def mouseMoveEvent(self, e):
        if self._drag:
            self.move(e.globalPosition().toPoint()-self._drag)

    def mouseReleaseEvent(self, e):
        self.pos_xy = [self.x(), self.y()]; self.save_data()

    # 持久化
    def export_data(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "导出数据（备份/换机用）", "quick_copy_data.json",
            "JSON 文件 (*.json)")
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump({"groups": self.groups,
                           "active_group": self.active_group,
                           "pos": [self.x(), self.y()],
                           "collapsed": self.collapsed,
                           "pin_on": self.pin_on, "masked": self.masked},
                          f, ensure_ascii=False, indent=2)
            self.flash(f"已导出到 {os.path.basename(path)}")
        except Exception as e:
            self.flash(f"导出失败：{e}", True)

    def import_data(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择要导入的 JSON", "", "JSON 文件 (*.json)")
        if not path:
            return
        try:
            d = json.load(open(path, encoding="utf-8"))
        except Exception as e:
            self.flash(f"文件无法解析：{e}", True); return
        new_groups = self._normalize_import(d)
        if not new_groups:
            self.flash("文件里没有可用条目", True); return
        mode, ok = QInputDialog.getItem(
            self, "导入方式", "选择导入方式：",
            ["替换全部数据", "按组合并（重名组/字段自动跳过）"], 0, False)
        if not ok:
            return
        if mode.startswith("替换"):
            self.groups = new_groups
            self.active_group = 0
        else:
            self._merge_groups(new_groups)
        self.refresh_groups(); self.refresh_rows(); self.save_data()
        self.flash("导入完成")

    @staticmethod
    def _normalize_import(d):
        """兼容 groups 版 / 旧平铺版 / 纯列表"""
        if isinstance(d, dict) and d.get("groups"):
            gs = []
            for g in d["groups"]:
                if isinstance(g, dict) and g.get("fields") is not None:
                    gs.append({"name": str(g.get("name") or "未命名"),
                               "fields": g["fields"]})
            return gs
        fields = None
        if isinstance(d, dict):
            fields = d.get("fields")
        elif isinstance(d, list):
            fields = d
        if isinstance(fields, list):
            return [{"name": "导入数据", "fields": fields}]
        return []

    def _merge_groups(self, new_groups):
        for ng in new_groups:
            hit = next((g for g in self.groups
                        if g["name"] == ng["name"]), None)
            if hit is None:
                self.groups.append({"name": ng["name"],
                                    "fields": list(ng["fields"])})
                continue
            exist = {f["name"] for f in hit["fields"]}
            for f in ng["fields"]:
                if f["name"] not in exist:
                    hit["fields"].append(f); exist.add(f["name"])

    def load_data(self):
        d = None
        if os.path.exists(DATA_FILE):
            try:
                d = json.load(open(DATA_FILE, encoding="utf-8"))
            except Exception:
                d = None
        if d and d.get("groups"):
            self.groups = d["groups"]
            self.active_group = min(d.get("active_group",0),len(self.groups)-1)
            self.pos_xy = d.get("pos"); self.collapsed = d.get("collapsed",False)
            self.pin_on = d.get("pin_on",True); self.masked = d.get("masked",False)
        elif d and d.get("fields"):
            self.groups = [{"name":"求职信息","fields":d["fields"]},
                           {"name":"应用密码","fields":[]}]
        else:
            self.groups = [{"name":"求职信息","fields":[]},
                           {"name":"应用密码","fields":[]}]

    def save_data(self):
        json.dump({"groups":self.groups,"active_group":self.active_group,
                   "pos":[self.x(),self.y()],"collapsed":self.collapsed,
                   "pin_on":self.pin_on,"masked":self.masked},
                  open(DATA_FILE,"w",encoding="utf-8"),
                  ensure_ascii=False, indent=2)

    def load_verify_cfg(self):
        if os.path.exists(CONFIG_FILE):
            try:
                self.verify_cfg.update(json.load(open(CONFIG_FILE,encoding="utf-8")))
            except Exception:
                pass

    def save_verify_cfg(self):
        json.dump(self.verify_cfg, open(CONFIG_FILE,"w",encoding="utf-8"),
                  ensure_ascii=False, indent=2)

    def close_win(self):
        self.save_data(); self.close()


def main():
    app = QApplication([])
    app.setStyleSheet(QSS)
    w = MainWindow(); w.show()
    app.exec()


if __name__ == "__main__":
    main()
