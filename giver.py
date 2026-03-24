"""
Giver Mail v1.1  —  @giver.com email client
Pure Python stdlib only. Works on Windows, macOS, Linux.
Python 3.12 / Windows DPI crash fixes applied.

Features:
  - @giver.com free local accounts  (users can message each other)
  - Connect any real IMAP account (Gmail, Outlook, Yahoo...)
  - Compose & send  (SMTP for real accounts, instant local for @giver.com)
  - Reply / Forward
  - Star, delete, mark read
  - Folders (Inbox / Sent / Drafts + IMAP folders)
  - Search
  - Import (JSON, .eml)  /  Export (JSON, CSV, .eml)
  - Settings  (SMTP config, accounts, import/export)
  - Beautiful background drawn by pure Python canvas art
  - Frosted-glass white panel over the background
  - OneDrive sync: auto-detects your OneDrive folder and stores all
    data there so Windows syncs it to the cloud automatically
"""

# ── Windows DPI fix — MUST come before tkinter ────────────────────────────────
import sys, os
if sys.platform == "win32":
    try:
        import ctypes
        try:    ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except: ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

import imaplib, smtplib, email, ssl, threading, json, hashlib, time, re
import csv, datetime, struct, zlib, base64, io, math
from email.header        import decode_header
from email.utils         import parseaddr, formatdate, make_msgid, parsedate_to_datetime
from email.mime.text     import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base     import MIMEBase
from email               import encoders

from tkinter import *
from tkinter import ttk, messagebox, filedialog
from tkinter.font import families as tk_families

# ── Show a dialog if anything crashes, instead of silent exit ─────────────────
def _crash(exc_type, exc_val, exc_tb):
    import traceback
    msg = "".join(traceback.format_exception(exc_type, exc_val, exc_tb))
    try:    messagebox.showerror("Giver crashed", msg)
    except: print(msg, file=sys.stderr)
sys.excepthook = _crash

# =============================================================================
#  STORAGE PATHS  —  OneDrive-aware
# =============================================================================

VERSION = "1.1"

def _find_onedrive():
    """
    Locate the local OneDrive sync folder on Windows.
    Returns the path if found, otherwise None.
    OneDrive automatically uploads anything saved here to the cloud.
    """
    if sys.platform != "win32":
        return None

    # 1. Check environment variables set by the OneDrive client
    for env_var in ("OneDriveCommercial", "OneDrive", "OneDriveConsumer"):
        path = os.environ.get(env_var, "")
        if path and os.path.isdir(path):
            return path

    # 2. Check the Windows registry (where OneDrive stores its path)
    try:
        import winreg
        key_paths = [
            r"SOFTWARE\Microsoft\OneDrive",
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders",
        ]
        for key_path in key_paths:
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                    for value_name in ("UserFolder", "OneDrive", "{374DE290-123F-4565-9164-39C4925E467B}"):
                        try:
                            val, _ = winreg.QueryValueEx(key, value_name)
                            if val and os.path.isdir(val):
                                return val
                        except FileNotFoundError:
                            pass
            except FileNotFoundError:
                pass
    except ImportError:
        pass

    # 3. Scan common folder names under the user home
    home = os.path.expanduser("~")
    for candidate in [
        "OneDrive",
        "OneDrive - Personal",
        "OneDrive - Telra",              # work/school OneDrive
        "OneDrive - Microsoft",
        "OneDriveConsumer",
    ]:
        full = os.path.join(home, candidate)
        if os.path.isdir(full):
            return full

    return None

def _choose_app_dir():
    """
    Pick where Giver Mail stores its data.
    Prefers OneDrive so files sync to the cloud automatically.
    Falls back to ~/.giver if OneDrive is not found.
    """
    onedrive = _find_onedrive()
    if onedrive:
        return os.path.join(onedrive, "GiverMail"), onedrive
    return os.path.join(os.path.expanduser("~"), ".giver"), None

APP_DIR, ONEDRIVE_ROOT = _choose_app_dir()
ACCS_FILE = os.path.join(APP_DIR, "accounts.json")
DB_FILE   = os.path.join(APP_DIR, "giver_db.json")
MAIL_DIR  = os.path.join(APP_DIR, "mailboxes")
SYNC_INFO = os.path.join(APP_DIR, "sync_info.json")   # records where data lives

os.makedirs(MAIL_DIR, exist_ok=True)

# Write a sync_info file so the user can see where data is stored
jsave_raw = lambda path, data: (
    os.makedirs(os.path.dirname(path), exist_ok=True) or
    open(path, "w", encoding="utf-8").write(
        __import__("json").dumps(data, indent=2, ensure_ascii=False))
)
try:
    jsave_raw(SYNC_INFO, {
        "app_dir":       APP_DIR,
        "onedrive_root": ONEDRIVE_ROOT,
        "synced_online": ONEDRIVE_ROOT is not None,
        "last_started":  time.strftime("%Y-%m-%d %H:%M:%S"),
    })
except Exception:
    pass

def jload(path, default=None):
    try:
        if os.path.exists(path):
            with open(path, encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return ({} if default is None else default)

def jsave(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

def _box_path(u, folder):
    return os.path.join(MAIL_DIR, u, f"{folder}.json")

def ensure_user(u):
    os.makedirs(os.path.join(MAIL_DIR, u), exist_ok=True)
    for box in ("inbox", "sent", "drafts"):
        p = _box_path(u, box)
        if not os.path.exists(p):
            jsave(p, [])

def load_box(u, folder="inbox"):
    return jload(_box_path(u, folder.lower()), [])

def save_box(u, folder, msgs):
    jsave(_box_path(u, folder.lower()), msgs)

def next_uid(u, folder):
    msgs = load_box(u, folder)
    ids  = [int(m["uid"]) for m in msgs if str(m.get("uid","")).isdigit()]
    return str(max(ids, default=0) + 1)

def mk_msg(uid, subject, from_a, to_a, body, date=None):
    now = time.strftime("%Y-%m-%d %H:%M:%S")
    return {
        "uid": uid, "subject": subject, "from": from_a, "to": to_a,
        "body": body, "read": False, "starred": False,
        "date": date or now, "date_fmt": time.strftime("%H:%M"),
    }

# =============================================================================
#  FONT RESOLVER  (safe on all platforms)
# =============================================================================

def resolve_fonts():
    avail = set(tk_families())
    def pick(*cands):
        for c in cands:
            if c in avail: return c
        return "TkDefaultFont"
    if sys.platform == "win32":
        base = pick("Segoe UI", "Tahoma", "Arial")
        mono = pick("Cascadia Code", "Consolas", "Courier New")
    elif sys.platform == "darwin":
        base = pick("SF Pro Text", "Helvetica Neue", "Helvetica", "Arial")
        mono = pick("SF Mono", "Menlo", "Monaco", "Courier")
    else:
        base = pick("Ubuntu", "DejaVu Sans", "Liberation Sans", "Helvetica", "Arial")
        mono = pick("Ubuntu Mono", "DejaVu Sans Mono", "Courier")
    return {
        "base": base, "mono": mono,
        "n":    (base, 10),
        "b":    (base, 10, "bold"),
        "s":    (base,  9),
        "sb":   (base,  9, "bold"),
        "t":    (base, 13, "bold"),
        "h":    (base, 11, "bold"),
        "big":  (base, 18, "bold"),
        "hero": (base, 22, "bold"),
        "body": (base, 10),
        "tiny": (base,  8),
    }

# =============================================================================
#  THEME
# =============================================================================

class T:
    # Core palette — teal/mint brand instead of Microsoft blue
    BRAND   = "#00897b"   # teal primary
    BRAND_H = "#00796b"
    BRAND_D = "#00695c"
    BRAND_L = "#e0f2f1"   # very light teal tint

    SIDE    = "#004d40"   # dark teal sidebar
    SIDE2   = "#00695c"
    SIDE_M  = "#80cbc4"

    # Glass panel colours
    GLASS   = "#f8fbff"   # main panel bg
    GLASS2  = "#edf4fb"   # secondary panel / hover
    BORDER  = "#c8dcea"

    # List
    SEL     = "#c8e6e0"
    HOV     = "#e8f5f2"
    UDOT    = "#00897b"

    # Text
    TEXT    = "#1a2630"
    TEXT2   = "#4a6070"
    TEXT3   = "#90a8b8"
    WHITE   = "#ffffff"

    # Semantic
    DANGER  = "#c62828"
    SUCCESS = "#1b5e20"
    WARN    = "#e65100"
    STAR    = "#f9a825"

    F = None   # filled after Tk()

# =============================================================================
#  BACKGROUND PAINTER  (pure Python Canvas art — no images needed)
# =============================================================================

def paint_background(canvas, W, H):
    """Draw a beautiful abstract background directly on a Canvas widget."""
    canvas.delete("all")

    # --- deep gradient via horizontal strips ---
    strips = 60
    for i in range(strips):
        t  = i / strips
        t2 = t * t
        r  = int(0   + 8   * t2)
        g  = int(30  + 60  * t2)
        b  = int(60  + 100 * t2)
        # clamp
        r = max(0, min(255, r))
        g = max(0, min(255, g))
        b = max(0, min(255, b))
        color = f"#{r:02x}{g:02x}{b:02x}"
        y0 = int(i     * H / strips)
        y1 = int((i+1) * H / strips) + 1
        canvas.create_rectangle(0, y0, W, y1, fill=color, outline="")

    # --- soft radial glow top-left ---
    cx, cy = W * 0.15, H * 0.18
    for radius in range(300, 0, -12):
        alpha = (1 - radius / 300) ** 2
        blue  = int(120 + alpha * 80)
        green = int(60  + alpha * 80)
        red   = int(0   + alpha * 20)
        blue  = max(0, min(255, blue))
        green = max(0, min(255, green))
        red   = max(0, min(255, red))
        col   = f"#{red:02x}{green:02x}{blue:02x}"
        canvas.create_oval(
            cx - radius, cy - radius, cx + radius, cy + radius,
            fill=col, outline="")

    # --- second glow bottom-right (teal tint) ---
    cx2, cy2 = W * 0.88, H * 0.80
    for radius in range(280, 0, -14):
        alpha = (1 - radius / 280) ** 2
        green = int(50 + alpha * 100)
        blue  = int(60 + alpha * 60)
        col   = f"#00{green:02x}{blue:02x}"
        canvas.create_oval(
            cx2 - radius, cy2 - radius, cx2 + radius, cy2 + radius,
            fill=col, outline="")

    # --- diagonal light streaks ---
    for i in range(8):
        x0 = -W * 0.3 + i * W * 0.18
        streak_w = W * 0.06
        pts = [x0, 0, x0 + streak_w, 0, x0 + streak_w + H * 0.4, H, x0 + H * 0.4, H]
        canvas.create_polygon(pts, fill="#ffffff", stipple="gray12", outline="")

    # --- subtle floating circles ---
    circles = [
        (W*0.08, H*0.75, 90,  "#00897b", "gray25"),
        (W*0.92, H*0.12, 70,  "#006064", "gray25"),
        (W*0.50, H*0.90, 50,  "#004d40", "gray25"),
        (W*0.72, H*0.35, 40,  "#00796b", "gray25"),
        (W*0.25, H*0.45, 30,  "#00897b", "gray25"),
    ]
    for cx, cy, r, fill, stip in circles:
        canvas.create_oval(cx-r, cy-r, cx+r, cy+r,
                           fill=fill, stipple=stip, outline="")

    # --- tiny star dots ---
    import random
    rng = random.Random(42)
    for _ in range(60):
        x = rng.randint(0, W)
        y = rng.randint(0, H)
        s = rng.choice([1, 1, 1, 2])
        br = rng.randint(80, 180)
        col = f"#{br:02x}{br:02x}{br:02x}"
        canvas.create_oval(x, y, x+s, y+s, fill=col, outline="")

# =============================================================================
#  SMALL REUSABLE WIDGETS
# =============================================================================

def hdiv(p, color=None, **kw):
    return Frame(p, bg=color or T.BORDER, height=1, **kw)

def vdiv(p, color=None, **kw):
    return Frame(p, bg=color or T.BORDER, width=1, **kw)

def flat_btn(parent, text, cmd, bg=None, fg=None, hov=None,
             font=None, px=14, py=7, side=None):
    bg  = bg  or T.BRAND
    fg  = fg  or T.WHITE
    hov = hov or T.BRAND_H
    fnt = font or T.F["b"]
    f   = Frame(parent, bg=bg, cursor="hand2")
    l   = Label(f, text=text, bg=bg, fg=fg, font=fnt, padx=px, pady=py)
    l.pack()
    def on_e(e): f.config(bg=hov); l.config(bg=hov)
    def on_l(e): f.config(bg=bg);  l.config(bg=bg)
    def on_c(e): cmd()
    for w in (f, l):
        w.bind("<Enter>",    on_e)
        w.bind("<Leave>",    on_l)
        w.bind("<Button-1>", on_c)
    if side: f.pack(side=side)
    return f

class NiceEntry(Frame):
    """Single-line entry with focus border highlight."""
    def __init__(self, parent, show=None, width=36, bg=None, font=None, **kw):
        bg = bg or T.WHITE
        super().__init__(parent, bg=T.BORDER, highlightthickness=0, **kw)
        inn = Frame(self, bg=bg)
        inn.pack(padx=1, pady=1)
        self.entry = Entry(inn, relief="flat", bg=bg, fg=T.TEXT,
                           insertbackground=T.BRAND, font=font or T.F["n"],
                           width=width, **({"show": show} if show else {}))
        self.entry.pack(padx=10, pady=7, fill=X)
        self.entry.bind("<FocusIn>",  lambda e: self.config(bg=T.BRAND))
        self.entry.bind("<FocusOut>", lambda e: self.config(bg=T.BORDER))

    def get(self):      return self.entry.get()
    def set(self, v):   self.entry.delete(0, END); self.entry.insert(0, v)
    def clear(self):    self.entry.delete(0, END)

class Avatar(Canvas):
    def __init__(self, parent, name="?", size=36, color=None, bg=None, **kw):
        pbg = bg or T.GLASS
        super().__init__(parent, width=size, height=size,
                         bg=pbg, highlightthickness=0, **kw)
        color = color or T.BRAND
        ini   = "".join(w[0].upper() for w in str(name).split()[:2]) or "?"
        r     = size // 2
        self.create_oval(1, 1, size-1, size-1, fill=color, outline="")
        self.create_text(r, r, text=ini, fill="white",
                         font=(T.F["base"], max(7, size // 3), "bold"))

# =============================================================================
#  EMAIL HELPERS
# =============================================================================

IMAP_MAP = {
    "gmail.com":      ("imap.gmail.com",            993),
    "googlemail.com": ("imap.gmail.com",            993),
    "outlook.com":    ("outlook.office365.com",     993),
    "hotmail.com":    ("outlook.office365.com",     993),
    "live.com":       ("outlook.office365.com",     993),
    "yahoo.com":      ("imap.mail.yahoo.com",       993),
    "ymail.com":      ("imap.mail.yahoo.com",       993),
    "icloud.com":     ("imap.mail.me.com",          993),
    "me.com":         ("imap.mail.me.com",          993),
    "telra.org":      ("outlook.office365.com",     993),
}
SMTP_MAP = {
    "gmail.com":      ("smtp.gmail.com",            587),
    "googlemail.com": ("smtp.gmail.com",            587),
    "outlook.com":    ("smtp.office365.com",        587),
    "hotmail.com":    ("smtp.office365.com",        587),
    "live.com":       ("smtp.office365.com",        587),
    "yahoo.com":      ("smtp.mail.yahoo.com",       587),
    "ymail.com":      ("smtp.mail.yahoo.com",       587),
    "icloud.com":     ("smtp.mail.me.com",          587),
    "telra.org":      ("smtp.office365.com",        587),
}

def dec(v):
    if not v: return ""
    return " ".join(
        p.decode(e or "utf-8", errors="replace") if isinstance(p, bytes) else p
        for p, e in decode_header(v)
    )

def fmt_date(s):
    if not s: return ""
    try:
        dt  = parsedate_to_datetime(s)
        now = datetime.datetime.now(dt.tzinfo)
        if dt.date() == now.date(): return dt.strftime("%H:%M")
        if (now - dt).days < 365:   return dt.strftime("%b %d")
        return dt.strftime("%b %d %Y")
    except:
        return s[:10]

def get_body(msg):
    plain = html = None
    if msg.is_multipart():
        for p in msg.walk():
            ct   = p.get_content_type()
            disp = str(p.get("Content-Disposition", ""))
            if "attachment" in disp: continue
            pl = p.get_payload(decode=True)
            if not pl: continue
            cs = p.get_content_charset() or "utf-8"
            if ct == "text/plain" and plain is None:
                plain = pl.decode(cs, errors="replace")
            elif ct == "text/html" and html is None:
                html  = pl.decode(cs, errors="replace")
    else:
        pl = msg.get_payload(decode=True)
        if pl:
            plain = pl.decode(msg.get_content_charset() or "utf-8", errors="replace")
    if plain: return plain
    if html:
        h = re.sub(r'<style[^>]*>.*?</style>', '', html, flags=re.DOTALL | re.I)
        h = re.sub(r'<[^>]+>', '', h)
        for a, b in [('&nbsp;', ' '), ('&amp;', '&'), ('&lt;', '<'),
                     ('&gt;', '>'), ('&quot;', '"'), ('&#39;', "'")]:
            h = h.replace(a, b)
        return re.sub(r'\n{3,}', '\n\n', h).strip()
    return "(No content)"

def get_attachments(msg):
    att = []
    if msg.is_multipart():
        for p in msg.walk():
            if "attachment" in str(p.get("Content-Disposition", "")):
                fn = p.get_filename()
                if fn:
                    att.append((dec(fn), p.get_payload(decode=True)))
    return att

# =============================================================================
#  IMAP CLIENT
# =============================================================================

class IMAPClient:
    def __init__(self): self.conn = None; self.current_folder = None

    def connect(self, host, port, user, pwd, use_ssl=True):
        if use_ssl:
            ctx = ssl.create_default_context()
            self.conn = imaplib.IMAP4_SSL(host, int(port), ssl_context=ctx)
        else:
            self.conn = imaplib.IMAP4(host, int(port))
        self.conn.login(user, pwd)

    def list_folders(self):
        if not self.conn: return []
        st, data = self.conn.list()
        result = []
        if st == "OK":
            for f in data:
                if isinstance(f, bytes): f = f.decode("utf-8", errors="replace")
                parts = f.split('"')
                name  = parts[-2] if len(parts) >= 3 else f.split()[-1]
                result.append(name.strip('/'))
        return result

    def select(self, folder="INBOX"):
        if not self.conn: return 0
        st, data = self.conn.select(f'"{folder}"')
        if st == "OK":
            self.current_folder = folder
            return int(data[0]) if data[0] else 0
        return 0

    def fetch_list(self, limit=100):
        if not self.conn: return []
        st, data = self.conn.search(None, "ALL")
        if st != "OK" or not data[0]: return []
        uids   = data[0].split()[-limit:]
        emails = []
        for uid in reversed(uids):
            st2, raw = self.conn.fetch(
                uid, "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE FLAGS)])")
            if st2 != "OK": continue
            if isinstance(raw[0], tuple):
                msg = email.message_from_bytes(raw[0][1])
                raw_bytes = raw[0][1]
                emails.append({
                    "uid":     uid.decode(),
                    "subject": dec(msg.get("Subject", "(No Subject)")),
                    "from":    dec(msg.get("From", "")),
                    "date":    fmt_date(msg.get("Date", "")),
                    "read":    b"\\Seen"    in raw_bytes,
                    "starred": b"\\Flagged" in raw_bytes,
                })
        return emails

    def fetch_full(self, uid):
        if not self.conn: return None
        st, raw = self.conn.fetch(uid.encode(), "(RFC822)")
        if st == "OK" and isinstance(raw[0], tuple):
            return email.message_from_bytes(raw[0][1])
        return None

    def mark_read(self, uid):
        self.conn.store(uid.encode(), "+FLAGS", "\\Seen")
    def delete(self, uid):
        self.conn.store(uid.encode(), "+FLAGS", "\\Deleted")
        self.conn.expunge()
    def star(self, uid):
        self.conn.store(uid.encode(), "+FLAGS", "\\Flagged")
    def unstar(self, uid):
        self.conn.store(uid.encode(), "-FLAGS", "\\Flagged")
    def disconnect(self):
        if self.conn:
            try: self.conn.close(); self.conn.logout()
            except: pass
            self.conn = None

# =============================================================================
#  SMTP SEND
# =============================================================================

def send_smtp(host, port, user, pwd, from_addr, to_list, subject, body,
              attachments=None):
    if isinstance(to_list, str):
        to_list = [a.strip() for a in re.split(r'[,;]', to_list) if a.strip()]
    msg = MIMEMultipart()
    msg["From"]       = from_addr
    msg["To"]         = ", ".join(to_list)
    msg["Subject"]    = subject
    msg["Date"]       = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid()
    msg.attach(MIMEText(body, "plain", "utf-8"))
    if attachments:
        for fp in attachments:
            with open(fp, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", "attachment",
                            filename=os.path.basename(fp))
            msg.attach(part)
    with smtplib.SMTP(host, int(port)) as s:
        s.starttls(context=ssl.create_default_context())
        s.login(user, pwd)
        s.sendmail(from_addr, to_list, msg.as_string())

def deliver_local(from_acc, to_email, subject, body):
    """Instant delivery between @giver.com accounts."""
    db      = jload(DB_FILE, {})
    to_user = to_email.split("@")[0].lower()
    if to_email.lower().endswith("@giver.com") and to_user in db:
        ensure_user(to_user)
        msgs = load_box(to_user, "inbox")
        msgs.insert(0, mk_msg(
            next_uid(to_user, "inbox"), subject,
            from_acc["email"], to_email, body))
        save_box(to_user, "inbox", msgs)
        return True
    return False

# =============================================================================
#  IMPORT / EXPORT
# =============================================================================

def export_json(msgs, path): jsave(path, msgs)

def export_eml(msgs, folder):
    os.makedirs(folder, exist_ok=True)
    for i, m in enumerate(msgs):
        fn = re.sub(r'[^\w\-_.]', '_', m.get("subject", "email"))[:40]
        with open(os.path.join(folder, f"{i+1:04d}_{fn}.eml"),
                  "w", encoding="utf-8") as f:
            f.write(f"From: {m.get('from','')}\n"
                    f"To: {m.get('to','')}\n"
                    f"Subject: {m.get('subject','')}\n"
                    f"Date: {m.get('date','')}\n\n"
                    f"{m.get('body','')}")

def export_csv(msgs, path):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["uid","subject","from","to","date","body","read"])
        w.writeheader()
        for m in msgs:
            w.writerow({k: m.get(k, "") for k in w.fieldnames})

def import_eml_file(path):
    with open(path, "rb") as f:
        msg = email.message_from_bytes(f.read())
    return mk_msg(
        str(int(time.time())),
        dec(msg.get("Subject", "")),
        dec(msg.get("From", "")),
        dec(msg.get("To", "")),
        get_body(msg),
        msg.get("Date", ""),
    )

# =============================================================================
#  COMPOSE WINDOW
# =============================================================================

class ComposeWin(Toplevel):
    def __init__(self, parent, account, reply_to=None, forward_of=None):
        super().__init__(parent)
        self.account       = account
        self.parent_win    = parent
        self._attachments  = []

        self.title("New Message — Giver Mail")
        self.configure(bg=T.GLASS)
        self.minsize(660, 500)
        self.grab_set()

        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        W, H   = min(740, sw - 80), min(580, sh - 80)
        self.geometry(f"{W}x{H}+{(sw-W)//2}+{(sh-H)//2}")

        self._build(reply_to, forward_of)

    def _build(self, reply_to, forward_of):
        # ── Toolbar ──────────────────────────────────────────────────────
        tb = Frame(self, bg=T.BRAND, height=46)
        tb.pack(fill=X)
        tb.pack_propagate(False)

        def tb_btn(text, cmd, fg=T.WHITE, side=LEFT):
            f = Frame(tb, bg=T.BRAND, cursor="hand2")
            f.pack(side=side, padx=4, pady=6)
            l = Label(f, text=text, bg=T.BRAND, fg=fg,
                      font=T.F["b"], padx=10, pady=5)
            l.pack()
            for w in (f, l):
                w.bind("<Enter>",    lambda e, b=f, lb=l: [b.config(bg=T.BRAND_H), lb.config(bg=T.BRAND_H)])
                w.bind("<Leave>",    lambda e, b=f, lb=l: [b.config(bg=T.BRAND),   lb.config(bg=T.BRAND)])
                w.bind("<Button-1>", lambda e, c=cmd: c())

        tb_btn("  Send  ",    self._send)
        tb_btn("  Attach ",   self._attach_file)
        tb_btn("  Draft  ",   self._save_draft)
        tb_btn("  Discard ",  self.destroy, fg="#ffcccc", side=RIGHT)

        # ── From strip ───────────────────────────────────────────────────
        fs = Frame(self, bg=T.GLASS2)
        fs.pack(fill=X)
        Label(fs, text=f"  From:   {self.account['email']}",
              bg=T.GLASS2, fg=T.TEXT2, font=T.F["s"]).pack(anchor=W, padx=8, pady=6)
        hdiv(self).pack(fill=X)

        # ── Address / Subject fields ──────────────────────────────────────
        def field_row(label):
            row = Frame(self, bg=T.WHITE)
            row.pack(fill=X)
            Label(row, text=label, bg=T.WHITE, fg=T.TEXT2,
                  font=T.F["s"], width=10, anchor=E).pack(side=LEFT, padx=(12, 0))
            e = Entry(row, relief="flat", bg=T.WHITE, fg=T.TEXT,
                      insertbackground=T.BRAND, font=T.F["n"], bd=0)
            e.pack(side=LEFT, fill=X, expand=True, padx=12, pady=8)
            hdiv(self).pack(fill=X)
            return e

        self._to_e   = field_row("To")
        self._cc_e   = field_row("Cc")
        self._subj_e = field_row("Subject")

        # Attachment bar (hidden until files added)
        self._att_bar = Frame(self, bg=T.BRAND_L)
        self._att_lbl = Label(self._att_bar, text="", bg=T.BRAND_L,
                              fg=T.BRAND_D, font=T.F["s"])
        self._att_lbl.pack(side=LEFT, padx=12, pady=4)

        # ── Body ─────────────────────────────────────────────────────────
        bw  = Frame(self, bg=T.WHITE)
        bw.pack(fill=BOTH, expand=True)
        bsb = ttk.Scrollbar(bw, orient=VERTICAL)
        self._body = Text(
            bw, bg=T.WHITE, fg=T.TEXT, insertbackground=T.BRAND,
            relief="flat", bd=0, font=T.F["body"], wrap=WORD,
            yscrollcommand=bsb.set, padx=16, pady=12,
            selectbackground=T.SEL)
        bsb.configure(command=self._body.yview)
        bsb.pack(side=RIGHT, fill=Y)
        self._body.pack(fill=BOTH, expand=True)

        # Pre-fill for reply / forward
        if reply_to:
            self._to_e.insert(0, reply_to.get("from", ""))
            s = reply_to.get("subject", "")
            self._subj_e.insert(0, s if s.startswith("Re:") else "Re: " + s)
            self._body.insert("1.0",
                f"\n\n─── Original message ───\n"
                f"From: {reply_to.get('from','')}\n"
                f"Date: {reply_to.get('date','')}\n\n"
                f"{reply_to.get('body','')}")
            self._body.mark_set("insert", "1.0")
        elif forward_of:
            s = forward_of.get("subject", "")
            self._subj_e.insert(0, s if s.startswith("Fwd:") else "Fwd: " + s)
            self._body.insert("1.0",
                f"\n\n─── Forwarded message ───\n"
                f"From: {forward_of.get('from','')}\n"
                f"To: {forward_of.get('to','')}\n"
                f"Date: {forward_of.get('date','')}\n\n"
                f"{forward_of.get('body','')}")
            self._body.mark_set("insert", "1.0")

        self._to_e.focus_set()

    def _attach_file(self):
        paths = filedialog.askopenfilenames(parent=self, title="Choose files to attach")
        if paths:
            self._attachments.extend(paths)
            names = ", ".join(os.path.basename(p) for p in self._attachments)
            self._att_lbl.config(text=f"  Attached: {names}")
            self._att_bar.pack(fill=X, before=self._body.master)

    def _save_draft(self):
        acc = self.account
        if acc.get("type") != "giver":
            messagebox.showinfo("Drafts",
                "Drafts are saved for @giver.com accounts only.", parent=self)
            return
        ensure_user(acc["username"])
        drafts = load_box(acc["username"], "drafts")
        drafts.insert(0, mk_msg(
            next_uid(acc["username"], "drafts"),
            self._subj_e.get(), acc["email"], self._to_e.get(),
            self._body.get("1.0", END).strip()))
        save_box(acc["username"], "drafts", drafts)
        messagebox.showinfo("Saved", "Draft saved.", parent=self)

    def _send(self):
        to      = self._to_e.get().strip()
        subj    = self._subj_e.get().strip()
        body    = self._body.get("1.0", END).strip()
        cc      = self._cc_e.get().strip()
        acc     = self.account

        if not to:
            messagebox.showerror("No recipient",
                "Please enter at least one recipient.", parent=self)
            return
        if not subj and not messagebox.askyesno(
                "No subject", "Send without a subject?", parent=self):
            return

        all_to = [a.strip() for a in re.split(r'[,;]', to + "," + cc) if a.strip()]

        # Split: @giver.com addresses get instant local delivery,
        # everything else goes via SMTP — regardless of account type.
        local_recips = [a for a in all_to if a.lower().endswith("@giver.com")]
        smtp_recips  = [a for a in all_to if not a.lower().endswith("@giver.com")]

        # ── Local delivery for @giver.com recipients ──────────────────
        local_ok, local_fail = [], []
        for addr in local_recips:
            (local_ok if deliver_local(acc, addr, subj, body)
             else local_fail).append(addr)

        # ── Save to Sent for @giver.com accounts ──────────────────────
        if acc.get("type") == "giver":
            ensure_user(acc["username"])
            sent = load_box(acc["username"], "sent")
            m    = mk_msg(next_uid(acc["username"], "sent"),
                          subj, acc["email"], to, body)
            m["read"] = True
            sent.insert(0, m)
            save_box(acc["username"], "sent", sent)

        # ── No SMTP recipients — done ─────────────────────────────────
        if not smtp_recips:
            if local_fail:
                messagebox.showwarning(
                    "Partial delivery",
                    f"Delivered to: {', '.join(local_ok) if local_ok else 'nobody'}\n"
                    f"No @giver.com account found for: {', '.join(local_fail)}",
                    parent=self)
            else:
                messagebox.showinfo("Sent",
                    f"Delivered to {', '.join(local_ok)}!", parent=self)
            if hasattr(self.parent_win, "_refresh_after_send"):
                self.parent_win._refresh_after_send()
            self.destroy()
            return

        # ── SMTP for all non-@giver.com recipients ────────────────────
        domain   = acc.get("email", "").split("@")[-1].lower()
        smtp_cfg = acc.get("smtp", {})
        s_host   = smtp_cfg.get("host") or SMTP_MAP.get(domain, ("", 587))[0]
        s_port   = smtp_cfg.get("port") or SMTP_MAP.get(domain, ("", 587))[1]
        s_user   = smtp_cfg.get("user") or acc.get("username", "")
        s_pass   = smtp_cfg.get("password") or acc.get("password", "")

        if not s_host:
            messagebox.showerror(
                "SMTP not configured",
                f"To send to {chr(10).join(smtp_recips)}, Giver Mail needs your "
                f"outgoing mail (SMTP) settings.\n\n"
                f"Go to  ⚙ Settings → SMTP / Send  and fill in:\n"
                f"  • SMTP Host  (e.g. smtp.gmail.com)\n"
                f"  • Port       (usually 587)\n"
                f"  • Username   (your full email address)\n"
                f"  • Password   (your password or App Password)\n\n"
                f"Gmail tip: use an App Password, not your normal password.\n"
                f"Get one at: myaccount.google.com → Security → App Passwords",
                parent=self)
            return

        self.config(cursor="watch"); self.update()

        def task():
            try:
                send_smtp(s_host, s_port, s_user, s_pass,
                          acc["email"], smtp_recips, subj, body,
                          attachments=self._attachments or None)
                parts = []
                if local_ok:    parts.append(f"Local: {', '.join(local_ok)}")
                if smtp_recips: parts.append(f"Email: {', '.join(smtp_recips)}")
                self.after(0, lambda: (
                    messagebox.showinfo("Sent!", "\n".join(parts), parent=self),
                    self.destroy()
                ))
                if hasattr(self.parent_win, "_refresh_after_send"):
                    self.after(0, self.parent_win._refresh_after_send)
            except Exception as ex:
                self.after(0, lambda: (
                    messagebox.showerror(
                        "Send failed",
                        f"SMTP error: {ex}\n\n"
                        f"Check your settings in  ⚙ Settings → SMTP / Send",
                        parent=self),
                    self.config(cursor="")
                ))

        threading.Thread(target=task, daemon=True).start()

# =============================================================================
#  SETTINGS WINDOW
# =============================================================================

class SettingsWin(Toplevel):
    def __init__(self, parent, account, on_change=None):
        super().__init__(parent)
        self.account   = account
        self.on_change = on_change
        self.title("Settings — Giver Mail")
        self.configure(bg=T.GLASS)
        self.minsize(620, 460)
        self.grab_set()

        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        W, H   = min(700, sw - 80), min(520, sh - 80)
        self.geometry(f"{W}x{H}+{(sw-W)//2}+{(sh-H)//2}")
        self._build()

    def _build(self):
        # Left nav
        nav = Frame(self, bg=T.GLASS2, width=165)
        nav.pack(side=LEFT, fill=Y)
        nav.pack_propagate(False)

        Label(nav, text="  Settings", bg=T.GLASS2, fg=T.TEXT,
              font=T.F["t"]).pack(anchor=W, padx=8, pady=(16, 8))
        hdiv(nav).pack(fill=X)

        self._pages     = {}
        self._nav_items = {}

        content = Frame(self, bg=T.GLASS)
        content.pack(side=LEFT, fill=BOTH, expand=True)
        self._content = content

        for name, builder in [
            ("Account",      self._pg_account),
            ("SMTP / Send",  self._pg_smtp),
            ("Import",       self._pg_import),
            ("Export",       self._pg_export),
            ("About",        self._pg_about),
        ]:
            page = Frame(content, bg=T.GLASS)
            self._pages[name] = page
            builder(page)

            f = Frame(nav, bg=T.GLASS2, cursor="hand2")
            f.pack(fill=X)
            l = Label(f, text=f"  {name}", bg=T.GLASS2, fg=T.TEXT2,
                      font=T.F["n"], anchor=W, padx=8, pady=9)
            l.pack(fill=X)
            self._nav_items[name] = (f, l)
            for w in (f, l):
                w.bind("<Button-1>", lambda e, n=name: self._show(n))
                w.bind("<Enter>",    lambda e, b=f, lb=l: [b.config(bg=T.HOV), lb.config(bg=T.HOV)])
                w.bind("<Leave>",    lambda e, b=f, lb=l, n=name:
                    [b.config(bg=T.SEL if self._cur==n else T.GLASS2),
                     lb.config(bg=T.SEL if self._cur==n else T.GLASS2)])

        self._cur = None
        self._show("Account")

    def _show(self, name):
        self._cur = name
        for n, (f, l) in self._nav_items.items():
            active = n == name
            f.config(bg=T.SEL     if active else T.GLASS2)
            l.config(bg=T.SEL     if active else T.GLASS2,
                     fg=T.BRAND   if active else T.TEXT2,
                     font=T.F["b"] if active else T.F["n"])
        for n, p in self._pages.items():
            p.pack_forget()
        self._pages[name].pack(fill=BOTH, expand=True, padx=26, pady=20)

    # ── Pages ─────────────────────────────────────────────────────────────

    def _pg_account(self, p):
        Label(p, text="Your account", bg=T.GLASS, fg=T.TEXT,
              font=T.F["big"]).pack(anchor=W)
        hdiv(p).pack(fill=X, pady=10)

        acc = self.account
        av  = Avatar(p, name=acc.get("name", "?"), size=52,
                     color=T.BRAND, bg=T.GLASS)
        av.pack(anchor=W, pady=(4, 10))

        rows = [
            ("Name",  acc.get("name", "")),
            ("Email", acc.get("email", "")),
            ("Type",  "Giver Mail (@giver.com)" if acc.get("type") == "giver"
                      else f"IMAP  ({acc.get('host','')})"),
        ]
        if acc.get("type") == "giver":
            rows.append(("Created", acc.get("created", "")))

        for label, value in rows:
            row = Frame(p, bg=T.GLASS)
            row.pack(fill=X, pady=2)
            Label(row, text=label + ":", bg=T.GLASS, fg=T.TEXT2,
                  font=T.F["s"], width=10, anchor=E).pack(side=LEFT)
            Label(row, text=str(value), bg=T.GLASS, fg=T.TEXT,
                  font=T.F["n"]).pack(side=LEFT, padx=8)

        hdiv(p).pack(fill=X, pady=14)

        Label(p, text="Saved accounts", bg=T.GLASS, fg=T.TEXT2,
              font=T.F["sb"]).pack(anchor=W, pady=(0, 6))

        accs = jload(ACCS_FILE, [])
        for a in accs:
            row = Frame(p, bg=T.GLASS2)
            row.pack(fill=X, pady=2)
            Label(row, text=a.get("email", ""), bg=T.GLASS2, fg=T.TEXT,
                  font=T.F["s"], padx=10, pady=5).pack(side=LEFT)
            rm_lbl = Label(row, text="Remove", bg=T.GLASS2, fg=T.DANGER,
                           font=T.F["s"], padx=8, cursor="hand2")
            rm_lbl.pack(side=RIGHT, pady=4)
            def rm(e=None, em=a["email"]):
                if messagebox.askyesno("Remove",
                        f"Remove {em} from saved accounts?", parent=self):
                    accs2 = [x for x in jload(ACCS_FILE, []) if x.get("email") != em]
                    jsave(ACCS_FILE, accs2)
                    messagebox.showinfo("Done",
                        "Removed. Restart to switch accounts.", parent=self)
            rm_lbl.bind("<Button-1>", rm)

    def _pg_smtp(self, p):
        Label(p, text="SMTP / Outgoing mail", bg=T.GLASS, fg=T.TEXT,
              font=T.F["big"]).pack(anchor=W)
        hdiv(p).pack(fill=X, pady=10)

        acc = self.account

        # Info banner — shown for giver accounts to explain hybrid sending
        if acc.get("type") == "giver":
            info = Frame(p, bg=T.BRAND_L, highlightthickness=1,
                         highlightbackground=T.BRAND)
            info.pack(fill=X, pady=(0, 14))
            Label(info,
                  text="  ✓  @giver.com addresses: delivered instantly, no SMTP needed.\n"
                       "  ✉  All other addresses (Gmail, Outlook…): fill in SMTP below.",
                  bg=T.BRAND_L, fg=T.BRAND_D,
                  font=T.F["s"], justify=LEFT).pack(anchor=W, padx=12, pady=10)

        domain   = acc.get("email", "").split("@")[-1].lower()
        smtp_cfg = acc.get("smtp", {})
        defs     = SMTP_MAP.get(domain, ("", 587))

        def row(lbl, default, show=None):
            Label(p, text=lbl, bg=T.GLASS, fg=T.TEXT2,
                  font=T.F["s"]).pack(anchor=W, pady=(8, 2))
            e = NiceEntry(p, show=show, bg=T.WHITE)
            e.set(str(default))
            e.pack(fill=X)
            return e

        self._sh = row("SMTP Host", smtp_cfg.get("host", defs[0]))
        self._sp = row("Port",      smtp_cfg.get("port", defs[1]))
        self._su = row("Username",  smtp_cfg.get("user", acc.get("username", "")))
        self._sw = row("Password",  smtp_cfg.get("password", acc.get("password", "")), show="*")

        Label(p, text="Most providers use port 587 with STARTTLS.",
              bg=T.GLASS, fg=T.TEXT3, font=T.F["s"]).pack(anchor=W, pady=(4, 12))

        def save_smtp():
            acc["smtp"] = {
                "host":     self._sh.get(),
                "port":     self._sp.get(),
                "user":     self._su.get(),
                "password": self._sw.get(),
            }
            accs = jload(ACCS_FILE, [])
            for i, a in enumerate(accs):
                if a.get("email") == acc.get("email"):
                    accs[i] = acc; break
            jsave(ACCS_FILE, accs)
            if self.on_change: self.on_change(acc)
            messagebox.showinfo("Saved", "SMTP settings saved.", parent=self)

        flat_btn(p, "Save", save_smtp, px=14, py=7).pack(anchor=W, pady=8)

    def _pg_import(self, p):
        Label(p, text="Import emails", bg=T.GLASS, fg=T.TEXT,
              font=T.F["big"]).pack(anchor=W)
        hdiv(p).pack(fill=X, pady=10)
        Label(p, text="Import messages into your @giver.com inbox.",
              bg=T.GLASS, fg=T.TEXT2, font=T.F["n"]).pack(anchor=W, pady=(0, 14))

        self._imp_status = Label(p, text="", bg=T.GLASS, fg=T.SUCCESS, font=T.F["s"])

        def do_import(msgs):
            acc = self.account
            if acc.get("type") != "giver":
                messagebox.showinfo("Info",
                    "Import adds to your @giver.com local inbox.", parent=self)
                return
            ensure_user(acc["username"])
            existing = load_box(acc["username"], "inbox")
            for m in msgs:
                m["uid"] = next_uid(acc["username"], "inbox")
                existing.insert(0, m)
            save_box(acc["username"], "inbox", existing)
            self._imp_status.config(text=f"Imported {len(msgs)} message(s).")

        def imp_json():
            path = filedialog.askopenfilename(
                parent=self, title="Choose JSON export",
                filetypes=[("JSON", "*.json"), ("All", "*.*")])
            if path:
                try:
                    msgs = jload(path, [])
                    if not isinstance(msgs, list): raise ValueError("Not a list")
                    do_import(msgs)
                except Exception as ex:
                    messagebox.showerror("Error", str(ex), parent=self)

        def imp_eml():
            paths = filedialog.askopenfilenames(
                parent=self, title="Choose .eml files",
                filetypes=[("EML", "*.eml"), ("All", "*.*")])
            if paths:
                msgs = []
                for path in paths:
                    try: msgs.append(import_eml_file(path))
                    except Exception as ex:
                        messagebox.showwarning("Skipped", str(ex), parent=self)
                if msgs: do_import(msgs)

        flat_btn(p, "Import from JSON backup", imp_json,
                 px=14, py=7).pack(anchor=W, pady=3)
        flat_btn(p, "Import .eml files",       imp_eml,
                 px=14, py=7).pack(anchor=W, pady=3)
        self._imp_status.pack(anchor=W, pady=6)

    def _pg_export(self, p):
        Label(p, text="Export emails", bg=T.GLASS, fg=T.TEXT,
              font=T.F["big"]).pack(anchor=W)
        hdiv(p).pack(fill=X, pady=10)
        Label(p, text="Export your local @giver.com emails for backup.",
              bg=T.GLASS, fg=T.TEXT2, font=T.F["n"]).pack(anchor=W, pady=(0, 10))

        Label(p, text="Folder", bg=T.GLASS, fg=T.TEXT2,
              font=T.F["sb"]).pack(anchor=W, pady=(0, 4))
        self._exp_folder = StringVar(value="inbox")
        fr = Frame(p, bg=T.GLASS)
        fr.pack(anchor=W, pady=(0, 12))
        for box in ("inbox", "sent", "drafts"):
            Radiobutton(fr, text=box.title(),
                        variable=self._exp_folder, value=box,
                        bg=T.GLASS, fg=T.TEXT,
                        activebackground=T.GLASS,
                        selectcolor=T.GLASS,
                        font=T.F["n"]).pack(side=LEFT, padx=4)

        self._exp_status = Label(p, text="", bg=T.GLASS, fg=T.SUCCESS, font=T.F["s"])

        def get_msgs():
            acc = self.account
            if acc.get("type") != "giver":
                messagebox.showinfo("Info",
                    "Export works with @giver.com accounts.", parent=self)
                return None
            return load_box(acc["username"], self._exp_folder.get())

        def do_json():
            msgs = get_msgs()
            if msgs is None: return
            path = filedialog.asksaveasfilename(
                parent=self, defaultextension=".json",
                filetypes=[("JSON", "*.json"), ("All", "*.*")])
            if path:
                export_json(msgs, path)
                self._exp_status.config(text=f"Exported {len(msgs)} messages.")

        def do_eml():
            msgs = get_msgs()
            if msgs is None: return
            folder = filedialog.askdirectory(parent=self, title="Save .eml files to")
            if folder:
                export_eml(msgs, folder)
                self._exp_status.config(text=f"Exported {len(msgs)} .eml files.")

        def do_csv():
            msgs = get_msgs()
            if msgs is None: return
            path = filedialog.asksaveasfilename(
                parent=self, defaultextension=".csv",
                filetypes=[("CSV", "*.csv"), ("All", "*.*")])
            if path:
                export_csv(msgs, path)
                self._exp_status.config(text=f"Exported {len(msgs)} rows to CSV.")

        flat_btn(p, "Export as JSON (backup)", do_json, px=14, py=7).pack(anchor=W, pady=3)
        flat_btn(p, "Export as .eml files",    do_eml,  px=14, py=7).pack(anchor=W, pady=3)
        flat_btn(p, "Export as CSV",            do_csv,  px=14, py=7).pack(anchor=W, pady=3)
        self._exp_status.pack(anchor=W, pady=8)

    def _pg_about(self, p):
        Label(p, text="Giver Mail", bg=T.GLASS, fg=T.TEXT,
              font=T.F["hero"]).pack(anchor=W)
        Label(p, text=f"v{VERSION}  —  Pure Python email client",
              bg=T.GLASS, fg=T.TEXT2, font=T.F["n"]).pack(anchor=W, pady=(4, 14))
        hdiv(p).pack(fill=X, pady=6)

        # ── Storage / sync status ─────────────────────────────────────────
        sync_frame = Frame(p, bg=T.BRAND_L, highlightthickness=1,
                           highlightbackground=T.BRAND)
        sync_frame.pack(fill=X, pady=(0, 14))

        if ONEDRIVE_ROOT:
            Label(sync_frame,
                  text="  ✓  Syncing to OneDrive",
                  bg=T.BRAND_L, fg=T.SUCCESS,
                  font=T.F["b"]).pack(anchor=W, padx=12, pady=(8, 2))
            Label(sync_frame,
                  text=f"  Data folder:  {APP_DIR}",
                  bg=T.BRAND_L, fg=T.TEXT2,
                  font=T.F["s"]).pack(anchor=W, padx=12)
            Label(sync_frame,
                  text=f"  OneDrive root: {ONEDRIVE_ROOT}",
                  bg=T.BRAND_L, fg=T.TEXT2,
                  font=T.F["s"]).pack(anchor=W, padx=12)
            Label(sync_frame,
                  text="  Windows will upload your mail automatically.",
                  bg=T.BRAND_L, fg=T.TEXT3,
                  font=T.F["s"]).pack(anchor=W, padx=12, pady=(0, 8))
        else:
            Label(sync_frame,
                  text="  ⚠  OneDrive not found — storing locally only",
                  bg="#fff8e1", fg=T.WARN,
                  font=T.F["b"]).pack(anchor=W, padx=12, pady=(8, 2))
            Label(sync_frame,
                  text=f"  Data folder:  {APP_DIR}",
                  bg="#fff8e1", fg=T.TEXT2,
                  font=T.F["s"]).pack(anchor=W, padx=12)
            Label(sync_frame,
                  text="  Install & sign in to OneDrive, then restart Giver Mail.",
                  bg="#fff8e1", fg=T.TEXT3,
                  font=T.F["s"]).pack(anchor=W, padx=12, pady=(0, 8))
            sync_frame.config(bg="#fff8e1", highlightbackground=T.WARN)

        def open_folder():
            try:
                if sys.platform == "win32":
                    os.startfile(APP_DIR)
                elif sys.platform == "darwin":
                    os.system(f'open "{APP_DIR}"')
                else:
                    os.system(f'xdg-open "{APP_DIR}"')
            except Exception:
                pass

        flat_btn(sync_frame, "Open data folder", open_folder,
                 bg=T.BRAND, px=12, py=5,
                 font=T.F["s"]).pack(anchor=W, padx=12, pady=(0, 10))

        hdiv(p).pack(fill=X, pady=6)

        for line in [
            ("Features:", T.TEXT, T.F["b"]),
            ("  @giver.com local accounts", T.TEXT2, T.F["n"]),
            ("  Local instant messaging (giver-to-giver)", T.TEXT2, T.F["n"]),
            ("  Connect real IMAP accounts", T.TEXT2, T.F["n"]),
            ("  Compose & send via SMTP", T.TEXT2, T.F["n"]),
            ("  Reply & Forward", T.TEXT2, T.F["n"]),
            ("  Import (JSON, EML)  /  Export (JSON, CSV, EML)", T.TEXT2, T.F["n"]),
            ("  OneDrive sync (auto-detected)", T.TEXT2, T.F["n"]),
            ("", T.TEXT2, T.F["n"]),
            ("Pure stdlib — zero dependencies.", T.TEXT3, T.F["s"]),
        ]:
            Label(p, text=line[0], bg=T.GLASS, fg=line[1],
                  font=line[2], anchor=W).pack(anchor=W, pady=1)

# =============================================================================
#  ONBOARDING WINDOW
# =============================================================================

class OnboardingWin(Tk):
    def __init__(self):
        super().__init__()
        T.F = resolve_fonts()

        self.title("Giver Mail")
        self.resizable(True, True)
        self.configure(bg="#000000")
        self.result = None

        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        # Fit to screen — leave some margin
        W = min(980, sw - 40)
        H = min(680, sh - 60)
        self.minsize(700, 500)
        self.geometry(f"{W}x{H}+{(sw-W)//2}+{(sh-H)//2}")

        self._build(W, H)
        self.bind("<Configure>", self._on_resize)
        self.protocol("WM_DELETE_WINDOW", self.destroy)

    def _on_resize(self, e):
        if e.widget is self:
            self._bg_canvas.config(width=e.width, height=e.height)
            self.after(50, lambda: paint_background(
                self._bg_canvas, e.width, e.height))

    def _build(self, W, H):
        # ── Full-window background canvas ─────────────────────────────────
        self._bg_canvas = Canvas(self, width=W, height=H,
                                 highlightthickness=0, bd=0)
        self._bg_canvas.place(x=0, y=0, relwidth=1, relheight=1)
        self.after(10, lambda: paint_background(self._bg_canvas, W, H))

        # ── Glass card — fills most of the window, centered with padding ──
        card = Frame(self, bg=T.GLASS, bd=0,
                     highlightthickness=1,
                     highlightbackground=T.BORDER)
        # Use place with relative sizing so it always fits
        card.place(relx=0.5, rely=0.5, anchor=CENTER,
                   relwidth=0.92, relheight=0.90)

        # ── Left branding panel inside card ───────────────────────────────
        left = Frame(card, bg=T.SIDE, width=290)
        left.pack(side=LEFT, fill=Y)
        left.pack_propagate(False)
        self._build_left(left)

        # ── Right content panel ───────────────────────────────────────────
        right = Frame(card, bg=T.GLASS)
        right.pack(side=LEFT, fill=BOTH, expand=True)
        self._build_right(right)

    def _build_left(self, p):
        mid = Frame(p, bg=T.SIDE)
        mid.place(relx=0.5, rely=0.44, anchor=CENTER)

        # G logo
        c = Canvas(mid, width=64, height=64, bg=T.SIDE, highlightthickness=0)
        c.pack(pady=(0, 16))
        c.create_oval(4, 4, 60, 60, fill=T.BRAND, outline=T.BRAND_L, width=2)
        c.create_text(32, 32, text="G", fill=T.WHITE,
                      font=(T.F["base"], 26, "bold"))

        Label(mid, text="Giver Mail", bg=T.SIDE, fg=T.WHITE,
              font=(T.F["base"], 22, "bold")).pack()
        Label(mid, text="Your free @giver.com inbox", bg=T.SIDE,
              fg=T.SIDE_M, font=T.F["s"]).pack(pady=(4, 0))

        hdiv(mid, color="#1a6b5e").pack(fill=X, pady=18)

        for feat in [
            "Free @giver.com address",
            "Send to any @giver.com user instantly",
            "Connect Gmail, Outlook, Yahoo…",
            "Compose, reply, forward",
            "Import & export your mail",
            "Settings & SMTP config",
        ]:
            row = Frame(mid, bg=T.SIDE)
            row.pack(fill=X, pady=2)
            Label(row, text="  ●  ", bg=T.SIDE, fg=T.BRAND,
                  font=(T.F["base"], 7)).pack(side=LEFT)
            Label(row, text=feat, bg=T.SIDE, fg=T.SIDE_M,
                  font=T.F["s"]).pack(side=LEFT)

    def _build_right(self, p):
        # Tab bar — 3 tabs
        tb = Frame(p, bg=T.GLASS, height=50)
        tb.pack(fill=X)
        tb.pack_propagate(False)

        self._tabs = {}
        tab_titles = ["Sign in", "Create account", "IMAP / Other"]
        for idx, title in enumerate(tab_titles):
            f = Frame(tb, bg=T.GLASS, cursor="hand2")
            f.pack(side=LEFT, fill=Y)
            l = Label(f, text=title, bg=T.GLASS, fg=T.TEXT2,
                      font=T.F["n"], padx=18, pady=14)
            l.pack()
            ab = Frame(f, height=3, bg=T.GLASS)
            ab.pack(fill=X, side=BOTTOM)
            self._tabs[idx] = {"f": f, "l": l, "ab": ab}
            for w in (f, l):
                w.bind("<Button-1>", lambda e, i=idx: self._switch(i))

        hdiv(p).pack(fill=X)

        # ── Scrollable page area ──────────────────────────────────────────
        scroll_wrap = Frame(p, bg=T.GLASS)
        scroll_wrap.pack(fill=BOTH, expand=True)

        self._page_canvas = Canvas(scroll_wrap, bg=T.GLASS,
                                   highlightthickness=0, bd=0)
        self._page_sb = ttk.Scrollbar(scroll_wrap, orient=VERTICAL,
                                      command=self._page_canvas.yview)
        self._page_canvas.configure(yscrollcommand=self._page_sb.set)
        self._page_sb.pack(side=RIGHT, fill=Y)
        self._page_canvas.pack(side=LEFT, fill=BOTH, expand=True)

        self._page_inner = Frame(self._page_canvas, bg=T.GLASS)
        self._page_win   = self._page_canvas.create_window(
            (0, 0), window=self._page_inner, anchor=NW)

        def _sync_width(e):
            self._page_canvas.itemconfig(self._page_win, width=e.width)
        self._page_canvas.bind("<Configure>", _sync_width)

        def _sync_scroll(e):
            self._page_canvas.configure(
                scrollregion=self._page_canvas.bbox("all"))
        self._page_inner.bind("<Configure>", _sync_scroll)

        def _on_wheel(e):
            delta = -1*(e.delta//120) if e.delta else (-1 if e.num==4 else 1)
            self._page_canvas.yview_scroll(delta, "units")
        for widget in (self._page_canvas, self._page_inner):
            widget.bind("<MouseWheel>", _on_wheel)
            widget.bind("<Button-4>",   _on_wheel)
            widget.bind("<Button-5>",   _on_wheel)

        area = Frame(self._page_inner, bg=T.GLASS)
        area.pack(fill=BOTH, expand=True, padx=36, pady=20)

        self._pg_signin = Frame(area, bg=T.GLASS)
        self._pg_new    = Frame(area, bg=T.GLASS)
        self._pg_imap   = Frame(area, bg=T.GLASS)

        self._build_pg_signin(self._pg_signin)
        self._build_pg_new(self._pg_new)
        self._build_pg_imap(self._pg_imap)

        # Auto-open Sign In if accounts exist, else Create
        existing = jload(DB_FILE, {})
        self._switch(0 if existing else 1)

    def _switch(self, idx):
        for i, t in self._tabs.items():
            a = i == idx
            t["l"].config(fg=T.BRAND  if a else T.TEXT2,
                          font=T.F["b"] if a else T.F["n"])
            t["ab"].config(bg=T.BRAND if a else T.GLASS)
        self._pg_signin.pack_forget()
        self._pg_new.pack_forget()
        self._pg_imap.pack_forget()
        page = {0: self._pg_signin, 1: self._pg_new, 2: self._pg_imap}[idx]
        page.pack(fill=BOTH, expand=True)
        # scroll to top
        self._page_canvas.yview_moveto(0)
        # Propagate mousewheel from every child widget to the scroll canvas
        self.after(50, lambda: self._bind_wheel_recursive(page))

    def _bind_wheel_recursive(self, widget):
        def _mw(e):
            delta = -1*(e.delta//120) if e.delta else (-1 if e.num==4 else 1)
            self._page_canvas.yview_scroll(delta, "units")
        for ev in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            try: widget.bind(ev, _mw)
            except Exception: pass
        for child in widget.winfo_children():
            self._bind_wheel_recursive(child)

    # ── Sign in to existing @giver.com account ──────────────────────────────

    def _build_pg_signin(self, p):
        Label(p, text="Sign in to Giver Mail", bg=T.GLASS, fg=T.TEXT,
              font=T.F["big"]).pack(anchor=W)
        Label(p, text="Sign in with your existing @giver.com account.",
              bg=T.GLASS, fg=T.TEXT2, font=T.F["s"]).pack(anchor=W, pady=(3, 18))

        # Show saved accounts as quick-pick buttons
        db = jload(DB_FILE, {})
        if db:
            Label(p, text="Your accounts", bg=T.GLASS, fg=T.TEXT2,
                  font=T.F["sb"]).pack(anchor=W, pady=(0, 8))

            for username, acc_data in db.items():
                self._make_account_card(p, acc_data)

            hdiv(p).pack(fill=X, pady=16)
            Label(p, text="Or sign in manually:", bg=T.GLASS,
                  fg=T.TEXT2, font=T.F["sb"]).pack(anchor=W, pady=(0, 8))

        # Manual sign-in fields
        Label(p, text="Username", bg=T.GLASS, fg=T.TEXT2,
              font=T.F["s"]).pack(anchor=W, pady=(0, 2))
        usr_row = Frame(p, bg=T.GLASS)
        usr_row.pack(fill=X)
        self._si_user = NiceEntry(usr_row, bg=T.WHITE)
        self._si_user.pack(side=LEFT, fill=X, expand=True)
        Label(usr_row, text="@giver.com", bg=T.GLASS,
              fg=T.TEXT2, font=T.F["n"]).pack(side=LEFT, padx=(6, 0))

        Label(p, text="Password", bg=T.GLASS, fg=T.TEXT2,
              font=T.F["s"]).pack(anchor=W, pady=(10, 2))
        self._si_pwd = NiceEntry(p, show="*", bg=T.WHITE)
        self._si_pwd.pack(fill=X)

        self._si_err = Label(p, text="", bg=T.GLASS, fg=T.DANGER, font=T.F["s"])
        self._si_err.pack(anchor=W, pady=(4, 0))

        flat_btn(p, "Sign In", self._do_signin,
                 px=14, py=7).pack(anchor=W, pady=(14, 0))

        # Link to create account
        link_row = Frame(p, bg=T.GLASS)
        link_row.pack(anchor=W, pady=(12, 0))
        Label(link_row, text="No account?  ", bg=T.GLASS,
              fg=T.TEXT2, font=T.F["s"]).pack(side=LEFT)
        lnk = Label(link_row, text="Create one →", bg=T.GLASS,
                    fg=T.BRAND, font=T.F["s"], cursor="hand2")
        lnk.pack(side=LEFT)
        lnk.bind("<Button-1>", lambda e: self._switch(1))

    def _make_account_card(self, parent, acc_data):
        """Clickable card for a saved @giver.com account."""
        card = Frame(parent, bg=T.GLASS2, highlightthickness=1,
                     highlightbackground=T.BORDER, cursor="hand2")
        card.pack(fill=X, pady=3)

        inner = Frame(card, bg=T.GLASS2, pady=8, padx=12)
        inner.pack(fill=X)

        av = Avatar(inner, name=acc_data.get("name", "?"),
                    size=34, color=T.BRAND, bg=T.GLASS2)
        av.pack(side=LEFT)

        info = Frame(inner, bg=T.GLASS2)
        info.pack(side=LEFT, padx=10)
        Label(info, text=acc_data.get("name", ""),
              bg=T.GLASS2, fg=T.TEXT, font=T.F["b"]).pack(anchor=W)
        Label(info, text=acc_data.get("email", ""),
              bg=T.GLASS2, fg=T.TEXT2, font=T.F["s"]).pack(anchor=W)

        arrow = Label(inner, text="→", bg=T.GLASS2,
                      fg=T.BRAND, font=T.F["h"])
        arrow.pack(side=RIGHT)

        def hover_on(e):
            for w in [card, inner, info, av, arrow] + list(info.winfo_children()):
                try: w.config(bg=T.SEL)
                except: pass
            card.config(highlightbackground=T.BRAND)

        def hover_off(e):
            for w in [card, inner, info, av, arrow] + list(info.winfo_children()):
                try: w.config(bg=T.GLASS2)
                except: pass
            card.config(highlightbackground=T.BORDER)

        def click(e, a=acc_data):
            # Ask for password to authenticate
            self._quick_signin(a)

        for w in [card, inner, info, av, arrow] + list(info.winfo_children()):
            try:
                w.bind("<Enter>",    hover_on)
                w.bind("<Leave>",    hover_off)
                w.bind("<Button-1>", click)
            except: pass

    def _quick_signin(self, acc_data):
        """Password prompt for one-click sign-in from account card."""
        dlg = Toplevel(self)
        dlg.title("Sign in")
        dlg.configure(bg=T.GLASS)
        dlg.resizable(False, False)
        dlg.grab_set()
        sw, sh = dlg.winfo_screenwidth(), dlg.winfo_screenheight()
        dlg.geometry(f"360x240+{(sw-360)//2}+{(sh-240)//2}")

        pad = Frame(dlg, bg=T.GLASS, padx=28, pady=24)
        pad.pack(fill=BOTH, expand=True)

        av = Avatar(pad, name=acc_data.get("name","?"),
                    size=42, color=T.BRAND, bg=T.GLASS)
        av.pack()
        Label(pad, text=acc_data.get("name",""), bg=T.GLASS,
              fg=T.TEXT, font=T.F["b"]).pack(pady=(6,0))
        Label(pad, text=acc_data.get("email",""), bg=T.GLASS,
              fg=T.TEXT2, font=T.F["s"]).pack(pady=(2,10))

        pw_entry = NiceEntry(pad, show="*", bg=T.WHITE)
        pw_entry.entry.insert(0, "")
        pw_entry.pack(fill=X)
        pw_entry.entry.insert(0, "")

        err_lbl = Label(pad, text="", bg=T.GLASS, fg=T.DANGER, font=T.F["s"])
        err_lbl.pack(anchor=W, pady=(4,0))

        def attempt(e=None):
            pw = pw_entry.get()
            hashed = hashlib.sha256(pw.encode()).hexdigest()
            if hashed == acc_data.get("pw_hash",""):
                dlg.destroy()
                self.result = acc_data
                self.destroy()
            else:
                err_lbl.config(text="Wrong password. Try again.")
                pw_entry.clear()

        pw_entry.entry.bind("<Return>", attempt)
        flat_btn(pad, "Sign In", attempt, px=14, py=7).pack(anchor=W, pady=(10,0))
        pw_entry.entry.focus_set()

    def _do_signin(self):
        """Manual username + password sign-in."""
        username = self._si_user.get().strip().lower()
        password = self._si_pwd.get()

        if not username:
            self._si_err.config(text="Please enter your username.")
            return
        if not password:
            self._si_err.config(text="Please enter your password.")
            return

        db = jload(DB_FILE, {})
        if username not in db:
            self._si_err.config(
                text=f"No account found for {username}@giver.com")
            return

        acc_data = db[username]
        hashed   = hashlib.sha256(password.encode()).hexdigest()
        if hashed != acc_data.get("pw_hash", ""):
            self._si_err.config(text="Wrong password. Please try again.")
            self._si_pwd.clear()
            return

        self.result = acc_data
        self.destroy()

    # ── New @giver.com account ────────────────────────────────────────────

    def _build_pg_new(self, p):
        Label(p, text="Create your account", bg=T.GLASS, fg=T.TEXT,
              font=T.F["big"]).pack(anchor=W)
        Label(p, text="Get a free @giver.com address in seconds.",
              bg=T.GLASS, fg=T.TEXT2, font=T.F["s"]).pack(anchor=W, pady=(3, 14))

        # Name row
        nr = Frame(p, bg=T.GLASS)
        nr.pack(fill=X, pady=2)
        lc = Frame(nr, bg=T.GLASS); lc.pack(side=LEFT, fill=X, expand=True, padx=(0, 8))
        rc = Frame(nr, bg=T.GLASS); rc.pack(side=LEFT, fill=X, expand=True)

        Label(lc, text="First name", bg=T.GLASS, fg=T.TEXT2,
              font=T.F["s"]).pack(anchor=W, pady=(0, 2))
        self._fn = NiceEntry(lc, bg=T.WHITE); self._fn.pack(fill=X)

        Label(rc, text="Last name",  bg=T.GLASS, fg=T.TEXT2,
              font=T.F["s"]).pack(anchor=W, pady=(0, 2))
        self._ln = NiceEntry(rc, bg=T.WHITE); self._ln.pack(fill=X)

        # Username
        Label(p, text="Email address", bg=T.GLASS, fg=T.TEXT2,
              font=T.F["s"]).pack(anchor=W, pady=(10, 2))
        ur = Frame(p, bg=T.GLASS); ur.pack(fill=X)
        self._un = NiceEntry(ur, bg=T.WHITE); self._un.pack(side=LEFT, fill=X, expand=True)
        Label(ur, text="@giver.com", bg=T.GLASS, fg=T.TEXT2,
              font=T.F["n"]).pack(side=LEFT, padx=(6, 0))
        self._un.entry.bind("<KeyRelease>", self._chk)
        self._un_st = Label(p, text="", bg=T.GLASS, fg=T.SUCCESS, font=T.F["s"])
        self._un_st.pack(anchor=W, pady=(2, 0))

        # Password
        Label(p, text="Password", bg=T.GLASS, fg=T.TEXT2,
              font=T.F["s"]).pack(anchor=W, pady=(8, 2))
        self._pw = NiceEntry(p, show="*", bg=T.WHITE); self._pw.pack(fill=X)
        Label(p, text="Confirm password", bg=T.GLASS, fg=T.TEXT2,
              font=T.F["s"]).pack(anchor=W, pady=(6, 2))
        self._pw2 = NiceEntry(p, show="*", bg=T.WHITE); self._pw2.pack(fill=X)

        self._new_err = Label(p, text="", bg=T.GLASS, fg=T.DANGER, font=T.F["s"])
        self._new_err.pack(anchor=W, pady=(4, 0))

        flat_btn(p, "Create Account", self._do_create,
                 px=14, py=7).pack(anchor=W, pady=(12, 0))

    def _chk(self, e=None):
        u  = self._un.get().strip().lower()
        db = jload(DB_FILE, {})
        if not u:
            self._un_st.config(text="")
        elif not re.match(r'^[a-z0-9._]{3,30}$', u):
            self._un_st.config(
                text="3–30 chars: a-z, 0-9, dot, underscore",
                fg=T.DANGER)
        elif u in db:
            self._un_st.config(
                text=f"  {u}@giver.com is already taken",
                fg=T.DANGER)
        else:
            self._un_st.config(
                text=f"  {u}@giver.com is available",
                fg=T.SUCCESS)

    def _do_create(self):
        fn  = self._fn.get().strip()
        ln  = self._ln.get().strip()
        un  = self._un.get().strip().lower()
        pw  = self._pw.get()
        pw2 = self._pw2.get()

        def err(m): self._new_err.config(text=m)

        if not fn or not ln:  return err("Please enter your full name.")
        if not re.match(r'^[a-z0-9._]{3,30}$', un):
            return err("Invalid username. Use 3–30 chars: a-z 0-9 . _")
        if len(pw) < 6:       return err("Password must be at least 6 characters.")
        if pw != pw2:         return err("Passwords do not match.")

        db = jload(DB_FILE, {})
        if un in db:          return err(f"{un}@giver.com is already taken.")

        acc = {
            "type":     "giver",
            "name":     f"{fn} {ln}",
            "email":    f"{un}@giver.com",
            "username": un,
            "pw_hash":  hashlib.sha256(pw.encode()).hexdigest(),
            "created":  time.strftime("%Y-%m-%d %H:%M:%S"),
        }
        db[un] = acc; jsave(DB_FILE, db)
        accs = jload(ACCS_FILE, []); accs.append(acc); jsave(ACCS_FILE, accs)
        ensure_user(un)

        # Seed welcome messages
        now = time.strftime("%Y-%m-%d %H:%M:%S")
        msgs = [
            mk_msg("1", "Welcome to Giver Mail!",
                   "Giver Team <hello@giver.com>", acc["email"],
                   f"Hi {fn},\n\nWelcome! Your address is {acc['email']}.\n\n"
                   "You can:\n"
                   "  • Compose and send to other @giver.com users instantly\n"
                   "  • Connect a real account (Gmail, Outlook, Yahoo…)\n"
                   "  • Import / export mail in Settings\n\n"
                   "— The Giver Team", now),
            mk_msg("2", "Getting started",
                   "Giver Tips <tips@giver.com>", acc["email"],
                   "Quick tips:\n\n"
                   "  • Click  + Compose  to write a new email\n"
                   "  • Send to any @giver.com address for instant delivery\n"
                   "  • Open  ⚙ Settings  to set up SMTP, import, or export\n"
                   "  • Search bar at the top filters your messages\n\n"
                   "Happy mailing!", now),
        ]
        msgs[1]["read"] = True
        save_box(un, "inbox", msgs)

        self._show_welcome(acc)

    def _show_welcome(self, acc):
        ov  = Frame(self, bg=T.GLASS)
        ov.place(relx=0, rely=0, relwidth=1, relheight=1)
        mid = Frame(ov, bg=T.GLASS)
        mid.place(relx=0.5, rely=0.5, anchor=CENTER)

        # Animated-style checkmark
        c = Canvas(mid, width=76, height=76, bg=T.GLASS, highlightthickness=0)
        c.pack(pady=(0, 18))
        c.create_oval(2, 2, 74, 74, fill=T.SUCCESS, outline="")
        c.create_line(20, 40, 34, 54, 56, 24,
                      fill=T.WHITE, width=5,
                      capstyle="round", joinstyle="round")

        Label(mid, text="Account created!", bg=T.GLASS, fg=T.TEXT,
              font=T.F["hero"]).pack()
        Label(mid, text=acc["email"],      bg=T.GLASS, fg=T.BRAND,
              font=T.F["h"]).pack(pady=(6, 2))
        Label(mid, text=acc["name"],       bg=T.GLASS, fg=T.TEXT2,
              font=T.F["n"]).pack()
        hdiv(mid).pack(fill=X, pady=20)
        def go(): self.result = acc; self.destroy()
        flat_btn(mid, "Open Giver Mail  →", go, px=18, py=9).pack()

    # ── Existing IMAP account ─────────────────────────────────────────────

    def _build_pg_imap(self, p):
        Label(p, text="Connect existing account", bg=T.GLASS, fg=T.TEXT,
              font=T.F["big"]).pack(anchor=W)
        Label(p, text="Gmail, Outlook, Yahoo or any IMAP server.",
              bg=T.GLASS, fg=T.TEXT2, font=T.F["s"]).pack(anchor=W, pady=(3, 12))

        # Provider shortcuts
        pr = Frame(p, bg=T.GLASS); pr.pack(anchor=W, pady=(0, 10))
        for name, host, port in [
            ("Gmail",   "imap.gmail.com",           "993"),
            ("Outlook", "outlook.office365.com",    "993"),
            ("Yahoo",   "imap.mail.yahoo.com",      "993"),
            ("iCloud",  "imap.mail.me.com",         "993"),
        ]:
            btn = Frame(pr, bg=T.GLASS, highlightthickness=1,
                        highlightbackground=T.BORDER, cursor="hand2")
            btn.pack(side=LEFT, padx=(0, 6))
            bl  = Label(btn, text=name, bg=T.GLASS, fg=T.TEXT2,
                        font=T.F["s"], padx=12, pady=6)
            bl.pack()
            def fill(e=None, h=host, po=port):
                self._imap_host.set(h); self._imap_port.set(po)
            for w in (btn, bl):
                w.bind("<Button-1>", fill)
                w.bind("<Enter>",    lambda e, b=btn, l=bl:
                    [b.config(highlightbackground=T.BRAND), l.config(fg=T.BRAND)])
                w.bind("<Leave>",    lambda e, b=btn, l=bl:
                    [b.config(highlightbackground=T.BORDER), l.config(fg=T.TEXT2)])

        Label(p, text="Email address", bg=T.GLASS, fg=T.TEXT2,
              font=T.F["s"]).pack(anchor=W, pady=(4, 2))
        self._imap_email = NiceEntry(p, bg=T.WHITE); self._imap_email.pack(fill=X)

        Label(p, text="Password", bg=T.GLASS, fg=T.TEXT2,
              font=T.F["s"]).pack(anchor=W, pady=(8, 2))
        self._imap_pwd = NiceEntry(p, show="*", bg=T.WHITE); self._imap_pwd.pack(fill=X)

        # Advanced toggle
        self._adv_open = False
        self._adv_lbl  = Label(p, text="▶  Advanced IMAP settings",
                               bg=T.GLASS, fg=T.BRAND,
                               font=T.F["s"], cursor="hand2")
        self._adv_lbl.pack(anchor=W, pady=(10, 0))
        self._adv_lbl.bind("<Button-1>", self._toggle_adv)

        self._adv_f = Frame(p, bg=T.GLASS)
        Label(self._adv_f, text="IMAP Host", bg=T.GLASS, fg=T.TEXT2,
              font=T.F["s"]).pack(anchor=W, pady=(6, 2))
        self._imap_host = NiceEntry(self._adv_f, bg=T.WHITE)
        self._imap_host.pack(fill=X)
        Label(self._adv_f, text="Port", bg=T.GLASS, fg=T.TEXT2,
              font=T.F["s"]).pack(anchor=W, pady=(6, 2))
        self._imap_port = NiceEntry(self._adv_f, width=10, bg=T.WHITE)
        self._imap_port.pack(anchor=W)

        self._imap_err = Label(p, text="", bg=T.GLASS, fg=T.DANGER, font=T.F["s"])
        self._imap_err.pack(anchor=W, pady=(8, 0))

        br = Frame(p, bg=T.GLASS); br.pack(anchor=W, pady=(12, 0))
        flat_btn(br, "Sign In", self._do_imap, px=14, py=7).pack(side=LEFT)
        self._spin = Label(br, text="", bg=T.GLASS, fg=T.TEXT2, font=T.F["s"])
        self._spin.pack(side=LEFT, padx=12)

    def _toggle_adv(self, e=None):
        self._adv_open = not self._adv_open
        self._adv_f.pack(fill=X) if self._adv_open else self._adv_f.pack_forget()
        self._adv_lbl.config(
            text=("▼" if self._adv_open else "▶") + "  Advanced IMAP settings")

    def _do_imap(self):
        eml  = self._imap_email.get().strip()
        pwd  = self._imap_pwd.get()
        host = self._imap_host.get().strip()
        port = self._imap_port.get().strip() or "993"

        if not eml or not pwd:
            self._imap_err.config(text="Please enter email and password.")
            return

        if not host:
            domain = eml.split("@")[-1].lower() if "@" in eml else ""
            if domain in IMAP_MAP:
                host, port = IMAP_MAP[domain]
                host = str(host); port = str(port)
            else:
                self._imap_err.config(
                    text="Unknown provider. Open Advanced settings and enter the host.")
                return

        self._imap_err.config(text="")
        self._spin.config(text="Connecting…"); self.update()

        def task():
            try:
                c = IMAPClient()
                c.connect(host, port, eml, pwd, use_ssl=True)
                c.disconnect()
                acc = {
                    "type":     "imap",
                    "name":     eml.split("@")[0].replace(".", " ").title(),
                    "email":    eml,
                    "host":     host,
                    "port":     port,
                    "username": eml,
                    "password": pwd,
                }
                accs = jload(ACCS_FILE, [])
                accs = [a for a in accs if a.get("email") != eml]
                accs.append(acc)
                jsave(ACCS_FILE, accs)
                self.after(0, lambda: (
                    setattr(self, "result", acc), self.destroy()))
            except Exception as ex:
                self.after(0, lambda: (
                    self._imap_err.config(text=f"Failed: {ex}"),
                    self._spin.config(text="")))

        threading.Thread(target=task, daemon=True).start()

# =============================================================================
#  MAIN MAIL WINDOW
# =============================================================================

class MailWin(Tk):
    def __init__(self, account):
        super().__init__()
        T.F = resolve_fonts()

        self.account  = account
        self.client   = IMAPClient()
        self.emails   = []
        self.cur_uid  = None
        self.cur_msg  = None   # email.message object (IMAP only)
        self.cur_em   = None   # local dict
        self._starred = set()
        self._cur_folder = "inbox"

        name = account.get("name", account["email"])
        self.title(f"Giver Mail — {account['email']}")
        self.configure(bg="#000000")
        self.minsize(900, 560)

        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        W = min(1300, sw - 40)
        H = min(800,  sh - 60)
        self.geometry(f"{W}x{H}+{(sw-W)//2}+{(sh-H)//2}")

        self._W = W; self._H = H
        self._setup_ttk()
        self._build(W, H)
        self.after(120, self._connect)
        self.protocol("WM_DELETE_WINDOW", self._quit)

    def _quit(self):
        self.client.disconnect(); self.destroy()

    def _setup_ttk(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure(".", background=T.GLASS, foreground=T.TEXT, font=T.F["n"])
        s.configure("Vertical.TScrollbar",
                    background=T.GLASS2, troughcolor=T.GLASS2,
                    arrowcolor=T.TEXT3, borderwidth=0, relief="flat", width=8)
        s.map("Vertical.TScrollbar", background=[("active", T.BORDER)])
        s.configure("TProgressbar",
                    background=T.BRAND, troughcolor=T.GLASS2,
                    borderwidth=0, relief="flat")

    def _build(self, W, H):
        # ── Background canvas (full window) ───────────────────────────────
        self._bg = Canvas(self, width=W, height=H,
                          highlightthickness=0, bd=0)
        self._bg.place(x=0, y=0, relwidth=1, relheight=1)
        self.after(10, lambda: paint_background(self._bg, W, H))

        # ── Ribbon ────────────────────────────────────────────────────────
        self._build_ribbon()

        # ── Progress bar ──────────────────────────────────────────────────
        self._prog = ttk.Progressbar(self, mode="indeterminate")
        self._prog.pack(fill=X)
        self._prog.pack_forget()

        # ── 3-column body ─────────────────────────────────────────────────
        body = Frame(self, bg="")   # transparent-ish
        body.pack(fill=BOTH, expand=True)
        body.configure(bg=T.SIDE)   # fallback

        self._build_sidebar(body)
        vdiv(body, color=T.SIDE2).pack(side=LEFT, fill=Y)
        self._build_list_pane(body)
        vdiv(body, color=T.BORDER).pack(side=LEFT, fill=Y)
        self._build_reading_pane(body)

        # ── Status bar ────────────────────────────────────────────────────
        sb = Frame(self, bg=T.SIDE, height=22)
        sb.pack(side=BOTTOM, fill=X)
        sb.pack_propagate(False)
        self._svar = StringVar(value="Ready")
        self._cvar = StringVar()
        Label(sb, textvariable=self._svar, bg=T.SIDE, fg=T.SIDE_M,
              font=T.F["tiny"]).pack(side=LEFT, padx=12)
        Label(sb, textvariable=self._cvar, bg=T.SIDE, fg=T.SIDE_M,
              font=T.F["tiny"]).pack(side=RIGHT, padx=12)

    def _build_ribbon(self):
        rb = Frame(self, bg=T.SIDE, height=50)
        rb.pack(fill=X)
        rb.pack_propagate(False)

        # Logo
        lw = Frame(rb, bg=T.SIDE, width=240)
        lw.pack(side=LEFT, fill=Y)
        lw.pack_propagate(False)

        c = Canvas(lw, width=28, height=28, bg=T.SIDE, highlightthickness=0)
        c.pack(side=LEFT, padx=(14, 6), pady=11)
        c.create_oval(1, 1, 27, 27, fill=T.BRAND, outline="")
        c.create_text(14, 14, text="G", fill=T.WHITE,
                      font=(T.F["base"], 12, "bold"))
        Label(lw, text="Giver Mail", bg=T.SIDE, fg=T.WHITE,
              font=(T.F["base"], 12, "bold")).pack(side=LEFT)

        # Search
        sw = Frame(rb, bg=T.SIDE)
        sw.pack(side=LEFT, fill=Y, expand=True, padx=24)
        si = Frame(sw, bg=T.WHITE)
        si.place(relx=0, rely=0.5, relwidth=1, anchor=W, height=30)
        Label(si, text="  Search", bg=T.WHITE, fg=T.TEXT3,
              font=T.F["s"]).pack(side=LEFT)
        self._sv = StringVar()
        se = Entry(si, textvariable=self._sv, bg=T.WHITE, fg=T.TEXT,
                   relief="flat", insertbackground=T.BRAND, font=T.F["n"])
        se.pack(side=LEFT, fill=X, expand=True, padx=4)
        se.bind("<Return>", self._do_search)

        # Settings gear
        gear = Label(rb, text=" ⚙ Settings ", bg=T.SIDE, fg=T.SIDE_M,
                     font=T.F["s"], cursor="hand2")
        gear.pack(side=RIGHT, padx=6)
        gear.bind("<Button-1>", lambda e: self._open_settings())
        gear.bind("<Enter>",    lambda e: gear.config(fg=T.WHITE))
        gear.bind("<Leave>",    lambda e: gear.config(fg=T.SIDE_M))

        # OneDrive sync indicator
        if ONEDRIVE_ROOT:
            sync_lbl = Label(rb, text="☁ OneDrive syncing",
                             bg=T.SIDE, fg="#80cbc4",
                             font=T.F["tiny"], cursor="hand2")
        else:
            sync_lbl = Label(rb, text="⚠ Local only",
                             bg=T.SIDE, fg="#ffcc80",
                             font=T.F["tiny"], cursor="hand2")
        sync_lbl.pack(side=RIGHT, padx=4)
        sync_lbl.bind("<Button-1>", lambda e: self._open_settings())

        # Account avatar chip
        af = Frame(rb, bg=T.SIDE)
        af.pack(side=RIGHT, padx=10, fill=Y)
        av = Avatar(af, name=self.account.get("name", "?"),
                    size=30, color=T.BRAND_D, bg=T.SIDE)
        av.pack(side=RIGHT, pady=10)
        Label(af, text=self.account["email"][:26], bg=T.SIDE,
              fg=T.SIDE_M, font=T.F["tiny"]).pack(side=RIGHT, padx=(0, 8))

    def _build_sidebar(self, parent):
        sb = Frame(parent, bg=T.SIDE, width=240)
        sb.pack(side=LEFT, fill=Y)
        sb.pack_propagate(False)

        # Compose button
        pad = Frame(sb, bg=T.SIDE, pady=12, padx=14)
        pad.pack(fill=X)
        flat_btn(pad, "+ Compose", self._compose,
                 bg=T.WHITE, fg=T.BRAND, hov=T.BRAND_L,
                 font=T.F["b"], px=12, py=7).pack(fill=X)

        hdiv(sb, color="#1a6b5e").pack(fill=X, padx=14, pady=6)

        Label(sb, text="  FOLDERS", bg=T.SIDE, fg=T.SIDE_M,
              font=(T.F["base"], 8, "bold")).pack(anchor=W, pady=(2, 4))

        self._folder_lb = Listbox(
            sb, bg=T.SIDE, fg=T.WHITE,
            selectbackground=T.SIDE2, selectforeground=T.WHITE,
            relief="flat", borderwidth=0, activestyle="none",
            font=T.F["n"], highlightthickness=0, selectborderwidth=0)
        self._folder_lb.pack(fill=BOTH, expand=True, padx=6, pady=2)
        self._folder_lb.bind("<<ListboxSelect>>", self._on_folder_click)

        hdiv(sb, color="#1a6b5e").pack(fill=X, padx=14, pady=6)
        flat_btn(sb, "⟳  Refresh", self._refresh,
                 bg=T.SIDE, fg=T.SIDE_M, hov=T.SIDE2,
                 font=T.F["s"], px=12, py=6).pack(fill=X, padx=14, pady=(0, 10))

    def _build_list_pane(self, parent):
        lf = Frame(parent, bg=T.GLASS, width=340)
        lf.pack(side=LEFT, fill=Y)
        lf.pack_propagate(False)

        hdr = Frame(lf, bg=T.GLASS, height=44)
        hdr.pack(fill=X); hdr.pack_propagate(False)
        self._ftitle = Label(hdr, text="Inbox", bg=T.GLASS,
                             fg=T.TEXT, font=T.F["t"])
        self._ftitle.pack(side=LEFT, padx=14, pady=10)
        self._lcnt = Label(hdr, text="", bg=T.GLASS, fg=T.TEXT3, font=T.F["s"])
        self._lcnt.pack(side=LEFT)
        hdiv(lf).pack(fill=X)

        wrap = Frame(lf, bg=T.GLASS); wrap.pack(fill=BOTH, expand=True)
        self._lc  = Canvas(wrap, bg=T.GLASS, highlightthickness=0)
        self._lsb = ttk.Scrollbar(wrap, orient=VERTICAL, command=self._lc.yview)
        self._lc.configure(yscrollcommand=self._lsb.set)
        self._lsb.pack(side=RIGHT, fill=Y)
        self._lc.pack(side=LEFT, fill=BOTH, expand=True)
        self._li  = Frame(self._lc, bg=T.GLASS)
        self._lw  = self._lc.create_window((0, 0), window=self._li, anchor=NW)
        self._li.bind("<Configure>",
            lambda e: self._lc.configure(scrollregion=self._lc.bbox("all")))
        self._lc.bind("<Configure>",
            lambda e: self._lc.itemconfig(self._lw, width=e.width))
        for ev in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self._lc.bind(ev, self._scroll_list)

    def _scroll_list(self, e):
        d = -1*(e.delta//120) if e.delta else (-1 if e.num == 4 else 1)
        self._lc.yview_scroll(d, "units")

    def _build_reading_pane(self, parent):
        rp = Frame(parent, bg=T.GLASS)
        rp.pack(side=LEFT, fill=BOTH, expand=True)

        # Action toolbar
        tb = Frame(rp, bg=T.GLASS, height=46)
        tb.pack(fill=X); tb.pack_propagate(False)
        self._star_lbl = None

        for text, cmd, fg in [
            ("Reply",     self._reply,     T.TEXT2),
            ("Forward",   self._forward,   T.TEXT2),
            ("Delete",    self._delete,    T.DANGER),
            ("Mark read", self._mark_read, T.TEXT2),
            ("☆ Star",   self._toggle_star, T.STAR),
        ]:
            f = Frame(tb, bg=T.GLASS, cursor="hand2"); f.pack(side=LEFT, fill=Y)
            l = Label(f, text=text, bg=T.GLASS, fg=fg,
                      font=T.F["s"], padx=12, pady=13)
            l.pack()
            if text == "☆ Star": self._star_lbl = l
            for w in (f, l):
                w.bind("<Enter>",    lambda e, b=f, lb=l: [b.config(bg=T.HOV), lb.config(bg=T.HOV)])
                w.bind("<Leave>",    lambda e, b=f, lb=l: [b.config(bg=T.GLASS), lb.config(bg=T.GLASS)])
                w.bind("<Button-1>", lambda e, c=cmd: c())

        hdiv(rp).pack(fill=X)

        # Email header
        self._ehdr = Frame(rp, bg=T.GLASS)
        self._ehdr.pack(fill=X, padx=28, pady=(18, 0))
        self._lbl_subj = Label(self._ehdr, text="", bg=T.GLASS, fg=T.TEXT,
                               font=(T.F["base"], 17, "bold"),
                               wraplength=560, justify=LEFT, anchor=W)
        self._lbl_subj.pack(anchor=W)

        mr = Frame(self._ehdr, bg=T.GLASS); mr.pack(fill=X, anchor=W, pady=(10, 0))
        self._av_hold = Frame(mr, bg=T.GLASS, width=40)
        self._av_hold.pack(side=LEFT); self._av_hold.pack_propagate(False)
        ic = Frame(mr, bg=T.GLASS); ic.pack(side=LEFT, padx=8, fill=X, expand=True)
        self._lbl_from = Label(ic, text="", bg=T.GLASS, fg=T.TEXT,
                               font=T.F["b"]); self._lbl_from.pack(anchor=W)
        self._lbl_to   = Label(ic, text="", bg=T.GLASS, fg=T.TEXT2,
                               font=T.F["s"]); self._lbl_to.pack(anchor=W)
        self._lbl_date = Label(mr, text="", bg=T.GLASS, fg=T.TEXT3,
                               font=T.F["s"]); self._lbl_date.pack(side=RIGHT)
        self._lbl_att  = Label(self._ehdr, text="", bg=T.GLASS, fg=T.BRAND,
                               font=T.F["s"], cursor="hand2")
        self._lbl_att.pack(anchor=W, pady=(6, 0))
        self._lbl_att.bind("<Button-1>", lambda e: self._save_att())
        hdiv(rp).pack(fill=X, padx=28, pady=12)

        # Body text
        bw  = Frame(rp, bg=T.GLASS); bw.pack(fill=BOTH, expand=True)
        bsb = ttk.Scrollbar(bw, orient=VERTICAL)
        self._body = Text(bw, bg=T.GLASS, fg=T.TEXT,
                          insertbackground=T.BRAND, relief="flat", bd=0,
                          font=T.F["body"], wrap=WORD, state=DISABLED,
                          yscrollcommand=bsb.set, padx=28, pady=4,
                          selectbackground=T.SEL, spacing1=2, spacing3=2)
        bsb.configure(command=self._body.yview)
        bsb.pack(side=RIGHT, fill=Y); self._body.pack(fill=BOTH, expand=True)

        # Empty state
        self._empty = Frame(rp, bg=T.GLASS)
        ec = Canvas(self._empty, width=60, height=48,
                    bg=T.GLASS, highlightthickness=0); ec.pack()
        ec.create_rectangle(4, 12, 56, 44, fill=T.GLASS2,
                             outline=T.BORDER, width=2)
        ec.create_polygon(4, 12, 30, 30, 56, 12, fill=T.GLASS2,
                          outline=T.BORDER, width=2)
        Label(self._empty, text="Select a message to read",
              bg=T.GLASS, fg=T.TEXT3,
              font=(T.F["base"], 11)).pack(pady=(10, 0))
        self._empty.place(relx=0.5, rely=0.5, anchor=CENTER)

    # ── Connection ─────────────────────────────────────────────────────────

    def _connect(self):
        acc = self.account
        if acc.get("type") == "giver":
            self._set_status("Giver Mail local account")
            self._populate_folders(["Inbox", "Sent", "Drafts"])
            self._load_local("inbox")
            return
        self._set_status("Connecting…"); self._show_prog(True)
        def task():
            try:
                self.client.connect(acc["host"], acc["port"],
                                    acc["username"], acc["password"])
                folders = self.client.list_folders()
                self.after(0, lambda: self._on_connected(folders))
            except Exception as ex:
                self.after(0, lambda: (
                    self._set_status(f"Connection failed: {ex}"),
                    self._show_prog(False)))
        threading.Thread(target=task, daemon=True).start()

    def _on_connected(self, folders):
        self._show_prog(False)
        self._set_status(f"Connected — {self.account['email']}")
        self._populate_folders(folders)
        for f in folders:
            if f.upper().strip("/") == "INBOX":
                self._load_imap_folder(f); break
        else:
            if folders: self._load_imap_folder(folders[0])

    def _populate_folders(self, folders):
        self._folders_raw = folders
        icons = {
            "inbox": "  Inbox",   "sent": "  Sent",
            "drafts": "  Drafts", "trash": "  Trash",
            "spam": "  Spam",     "junk": "  Junk",
            "archive": "  Archive",
        }
        self._folder_lb.delete(0, END)
        for f in folders:
            key = f.lower().split("/")[-1].strip()
            self._folder_lb.insert(END, icons.get(key, f"  {f}"))

    def _on_folder_click(self, e):
        sel = self._folder_lb.curselection()
        if not sel: return
        folder = self._folders_raw[sel[0]]
        self._ftitle.config(text=folder)
        self._cur_folder = folder.lower()
        if self.account.get("type") == "giver":
            self._load_local(folder.lower())
        else:
            self._load_imap_folder(folder)

    def _load_local(self, folder="inbox"):
        acc = self.account
        ensure_user(acc["username"])
        msgs = load_box(acc["username"], folder)
        self._cur_folder = folder
        self._ftitle.config(text=folder.title())
        self._show_list(msgs, len(msgs))
        self._set_status(folder.title())

    def _load_imap_folder(self, folder):
        self._set_status(f"Loading {folder}…"); self._show_prog(True)
        def task():
            try:
                count  = self.client.select(folder)
                emails = self.client.fetch_list(limit=100)
                self.after(0, lambda: self._show_list(emails, count))
            except Exception as ex:
                self.after(0, lambda: (
                    self._set_status(f"Error: {ex}"),
                    self._show_prog(False)))
        threading.Thread(target=task, daemon=True).start()

    def _show_list(self, emails, count):
        self._show_prog(False)
        self.emails = emails
        self._lcnt.config(text=str(count))
        self._cvar.set(f"{len(emails)} messages")
        self._render_list(emails)

    def _render_list(self, emails):
        for w in self._li.winfo_children():
            w.destroy()
        for em in emails:
            self._make_row(em)
        self._lc.yview_moveto(0)

    def _make_row(self, em):
        is_read = em.get("read", True)
        _, addr  = parseaddr(em.get("from", ""))
        disp     = em.get("from", "").split("<")[0].strip().strip('"') or addr
        NORM = T.GLASS; HOV = T.HOV

        row   = Frame(self._li, bg=NORM, cursor="hand2"); row.pack(fill=X)
        dot   = Frame(row, bg=T.UDOT if not is_read else NORM, width=4)
        dot.pack(side=LEFT, fill=Y)
        inner = Frame(row, bg=NORM, padx=12, pady=9)
        inner.pack(side=LEFT, fill=BOTH, expand=True)
        top   = Frame(inner, bg=NORM); top.pack(fill=X)

        Label(top, text=disp[:30], bg=NORM, fg=T.TEXT,
              font=T.F["b"] if not is_read else T.F["n"]).pack(side=LEFT)
        Label(top, text=em.get("date", ""), bg=NORM, fg=T.TEXT3,
              font=T.F["tiny"]).pack(side=RIGHT)
        Label(inner, text=em.get("subject", "")[:50], bg=NORM,
              fg=T.TEXT if not is_read else T.TEXT2,
              font=T.F["sb"] if not is_read else T.F["s"]).pack(anchor=W)
        if em.get("starred"):
            Label(inner, text="★", bg=NORM, fg=T.STAR,
                  font=T.F["tiny"]).pack(anchor=W)
        hdiv(self._li).pack(fill=X)

        all_w = [row, dot, inner, top] + \
                list(inner.winfo_children()) + list(top.winfo_children())

        def on_e(e, ws=all_w, d=dot, rd=is_read):
            for w in ws:
                try: w.config(bg=HOV)
                except: pass
            if not rd: d.config(bg=T.UDOT)

        def on_l(e, ws=all_w, d=dot, rd=is_read):
            for w in ws:
                try: w.config(bg=NORM)
                except: pass
            d.config(bg=T.UDOT if not rd else NORM)

        def on_c(e, u=em["uid"], m=em):
            self._open_email(u, m)

        for w in all_w:
            try:
                w.bind("<Enter>",    on_e)
                w.bind("<Leave>",    on_l)
                w.bind("<Button-1>", on_c)
            except: pass

    # ── Reading pane ───────────────────────────────────────────────────────

    def _open_email(self, uid, em):
        self.cur_uid = uid; self.cur_em = em
        if em.get("body"):
            self.cur_msg = None; self._display(em)
        else:
            self._show_prog(True)
            def task():
                try:
                    msg = self.client.fetch_full(uid)
                    self.after(0, lambda: self._on_fetched(uid, msg))
                except Exception as ex:
                    self.after(0, lambda: (
                        self._set_status(f"Error: {ex}"),
                        self._show_prog(False)))
            threading.Thread(target=task, daemon=True).start()

    def _on_fetched(self, uid, msg):
        self._show_prog(False)
        if not msg: return
        self.cur_msg = msg
        self._display({
            "uid":         uid,
            "subject":     dec(msg.get("Subject", "")),
            "from":        dec(msg.get("From", "")),
            "to":          dec(msg.get("To", "")),
            "date":        msg.get("Date", ""),
            "body":        get_body(msg),
            "attachments": get_attachments(msg),
        })

    def _display(self, em):
        self._empty.place_forget()
        _, addr = parseaddr(em.get("from", ""))
        disp    = em.get("from", "").split("<")[0].strip().strip('"') or addr

        self._lbl_subj.config(text=em.get("subject", "") or "(No Subject)")
        self._lbl_from.config(text=disp)
        self._lbl_to.config(
            text=f"To: {em.get('to','')}" if em.get("to") else "")

        for w in self._av_hold.winfo_children(): w.destroy()
        Avatar(self._av_hold, name=disp, size=36,
               color=T.BRAND, bg=T.GLASS).pack()

        try:
            dt = parsedate_to_datetime(em.get("date", ""))
            ds = dt.strftime("%a %b %d, %Y  %H:%M")
        except:
            ds = em.get("date", "")[:24]
        self._lbl_date.config(text=ds)

        atts = em.get("attachments", [])
        self._lbl_att.config(
            text=f"  {len(atts)} attachment(s): "
                 f"{', '.join(n for n,_ in atts[:3])}  — click to save"
            if atts else "")

        self._body.config(state=NORMAL)
        self._body.delete("1.0", END)
        self._body.insert(END, em.get("body", ""))
        self._body.config(state=DISABLED)
        self._body.yview_moveto(0)

        starred = em.get("starred", False) or self.cur_uid in self._starred
        if self._star_lbl:
            self._star_lbl.config(
                text="★ Starred" if starred else "☆ Star",
                fg=T.STAR)

    # ── Toolbar actions ────────────────────────────────────────────────────

    def _compose(self):
        ComposeWin(self, self.account)

    def _reply(self):
        if self.cur_em: ComposeWin(self, self.account, reply_to=self.cur_em)

    def _forward(self):
        if self.cur_em: ComposeWin(self, self.account, forward_of=self.cur_em)

    def _refresh(self):
        if self.account.get("type") == "giver":
            self._load_local(self._cur_folder)
        elif self.client.current_folder:
            self._load_imap_folder(self.client.current_folder)

    def _refresh_after_send(self):
        if self._cur_folder in ("sent", "inbox"):
            self._load_local(self._cur_folder)

    def _mark_read(self):
        if not self.cur_uid: return
        if self.account.get("type") == "giver":
            acc  = self.account
            msgs = load_box(acc["username"], self._cur_folder)
            for m in msgs:
                if m["uid"] == self.cur_uid: m["read"] = True
            save_box(acc["username"], self._cur_folder, msgs)
            self._refresh()
        elif self.client.conn:
            try: self.client.mark_read(self.cur_uid)
            except: pass

    def _delete(self):
        if not self.cur_uid: return
        if not messagebox.askyesno("Delete", "Delete this email?", parent=self):
            return
        if self.account.get("type") == "giver":
            acc  = self.account
            msgs = load_box(acc["username"], self._cur_folder)
            msgs = [m for m in msgs if m["uid"] != self.cur_uid]
            save_box(acc["username"], self._cur_folder, msgs)
            self._clear_read(); self._refresh()
        elif self.client.conn:
            try:
                self.client.delete(self.cur_uid)
                self._clear_read(); self._refresh()
            except Exception as ex:
                self._set_status(f"Error: {ex}")

    def _toggle_star(self):
        if not self.cur_uid: return
        starred = self.cur_uid in self._starred
        if starred:
            self._starred.discard(self.cur_uid)
        else:
            self._starred.add(self.cur_uid)
        if self._star_lbl:
            self._star_lbl.config(
                text="★ Starred" if not starred else "☆ Star",
                fg=T.STAR)
        if self.account.get("type") == "giver" and self.cur_em:
            acc  = self.account
            msgs = load_box(acc["username"], self._cur_folder)
            for m in msgs:
                if m["uid"] == self.cur_uid:
                    m["starred"] = not starred
            save_box(acc["username"], self._cur_folder, msgs)

    def _save_att(self):
        if not self.cur_msg:
            messagebox.showinfo("No email",
                "Open an IMAP email to save attachments.", parent=self)
            return
        atts = get_attachments(self.cur_msg)
        if not atts:
            messagebox.showinfo("No attachments",
                "This email has no attachments.", parent=self)
            return
        folder = filedialog.askdirectory(
            title="Save attachments to…", parent=self)
        if not folder: return
        for fn, data in atts:
            with open(os.path.join(folder, fn), "wb") as f:
                f.write(data)
        messagebox.showinfo("Saved",
            f"Saved {len(atts)} file(s) to:\n{folder}", parent=self)

    def _do_search(self, e=None):
        q = self._sv.get().strip().lower()
        if not q:
            self._render_list(self.emails); return
        results = [m for m in self.emails
                   if q in m.get("subject", "").lower()
                   or q in m.get("from", "").lower()]
        self._render_list(results)
        self._lcnt.config(text=str(len(results)))

    def _open_settings(self):
        SettingsWin(self, self.account,
                    on_change=lambda a: setattr(self, "account", a))

    def _clear_read(self):
        for l in (self._lbl_subj, self._lbl_from, self._lbl_to,
                  self._lbl_date, self._lbl_att):
            l.config(text="")
        self._body.config(state=NORMAL)
        self._body.delete("1.0", END)
        self._body.config(state=DISABLED)
        self.cur_uid = None; self.cur_msg = None; self.cur_em = None
        self._empty.place(relx=0.5, rely=0.5, anchor=CENTER)

    def _show_prog(self, show):
        if show: self._prog.pack(fill=X); self._prog.start(8)
        else:    self._prog.stop(); self._prog.pack_forget()

    def _set_status(self, msg): self._svar.set(msg)

# =============================================================================
#  ENTRY POINT
# =============================================================================

def main():
    ob = OnboardingWin()
    ob.mainloop()
    if ob.result:
        app = MailWin(ob.result)
        app.mainloop()

if __name__ == "__main__":
    main()
