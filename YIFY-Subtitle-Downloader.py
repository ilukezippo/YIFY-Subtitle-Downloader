# YIFY-Subtitle-Downloader v1.1
# GUI subtitle fetcher for YIFY/YTS-Subs with IMDb matching (EN/AR)
# v1.1 adds:
# - Donate pill button next to signature
# - Checkbox images (Include column), Select All/None
# - Include column locked, Excel-like auto-fit on header right-edge double-click
# - Progress helpers (progress_start/step/finish)
# - Remembers last folder across launches
# - Window icon from icon.ico

import os
import re
import time
import json
import threading
import shutil
import urllib.parse
import tempfile
from zipfile import ZipFile
from difflib import SequenceMatcher
from io import BytesIO

import requests
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.font as tkfont
from PIL import Image, ImageDraw

# ----------------------------- Config -----------------------------
MIRRORS = [
    "https://yifysubtitles.ch",
    "https://yts-subs.com",
]

VIDEO_EXTS = {".mp4", ".mkv", ".avi", ".mov", ".wmv", ".flv", ".m4v"}
LANG_TOKENS = {"english": "en", "arabic": "ar"}      # site language -> filename suffix
MAP_FILENAME = ".yify_imdb_map.json"                 # per-folder mapping

# App state file (store last folder)
def _state_file() -> str:
    base = os.getenv("LOCALAPPDATA") or os.path.expanduser("~")
    d = os.path.join(base, "YIFYSubtitleDownloader")
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, "settings.json")

APP_STATE_FILE = _state_file()

# Networking
REQUEST_TIMEOUT = 10
RETRIES_PER_HOST = 2
BACKOFF_SECS = 0.8
BLOCK_CODES = {403, 429, 437}

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/129.0.0.0 Safari/537.36"),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Connection": "keep-alive",
})

# ----------------------------- Helpers: name cleaning -----------------------------
JUNK = (
    "480p|720p|1080p|2160p|4k|10bit|8bit|x264|x265|h\\.?264|h\\.?265|hevc|av1|"
    "webrip|web-dl|webdl|bluray|b[dr]rip|hdrip|dvdrip|cam|telesync|ts|r5|"
    "aac(?:\\d(?:\\.\\d)?)?|ac3|eac3|ddp(?:\\.\\d)?|dts(?:-hd)?|truehd|"
    "atmos|dolby|vision|hdr|hdr10\\+?|sdr|"
    "yts|yify|rarbg|ettv|evo|fgt|psa|pahe|tigole|ntb|vtv|xvid|proper|repack|remux|"
    "nf|amzn|hulu|dsnp|web|blu-ray|webrip"
)
STOPWORDS = {"the", "a", "an", "and", "or", "of", "in", "on", "at"}

def collapse(s: str) -> str:
    s = s.replace("_", " ").replace(".", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def normalize_title(s: str) -> str:
    s = s.lower()
    s = re.sub(r"[\[\]\(\)\{\}:;,'\"!@#$%^&*+=/?\\|~`]", " ", s)
    s = re.sub(r"\s+", " ", s)
    s = " ".join([w for w in s.split() if w not in STOPWORDS])
    return s.strip()

def filename_to_title_and_year(name: str):
    base, _ = os.path.splitext(name)
    s = collapse(base)

    m_end = re.search(r"\(((19|20)\d{2})\)\s*$", s)
    if m_end:
        year = m_end.group(1)
        title_segment = s[:m_end.start()]
    else:
        matches = list(re.finditer(r"\b(19|20)\d{2}\b", s))
        if matches:
            m = matches[-1]
            year = m.group(0)
            title_segment = s[:m.start()]
        else:
            year = None
            title_segment = s

    title = re.sub(r"[\[\(\{][^\]\)\}]*[\]\)\}]", " ", title_segment)
    title = re.sub(rf"(?i)\b(?:{JUNK})\b", " ", title)
    title = collapse(title)
    title = re.sub(r"[\s\-\–:\(\[\{]+$", "", title).strip()
    if not title and title_segment:
        title = collapse(title_segment)
        title = re.sub(r"[\s\-\–:\(\[\{]+$", "", title).strip()
    return title, year

# ----------------------------- Robust HTTP -----------------------------
def get_html(url, *, params=None, stream=False, headers=None):
    last_err = None
    for attempt in range(1 + RETRIES_PER_HOST):
        try:
            r = SESSION.get(url, params=params, timeout=REQUEST_TIMEOUT,
                            allow_redirects=True, stream=stream, headers=headers)
            if r.status_code in BLOCK_CODES:
                raise requests.HTTPError(f"Blocked {r.status_code}", response=r)
            r.raise_for_status()
            return r
        except Exception as e:
            last_err = e
            if attempt < RETRIES_PER_HOST:
                time.sleep(BACKOFF_SECS * (attempt + 1))
            else:
                break
    raise last_err

# ----------------------------- IMDb helpers -----------------------------
def imdb_suggest(title: str):
    q = title.strip()
    if not q:
        return []
    prefix = q[0].lower() if q[0].isalnum() else "1"
    url = f"https://v2.sg.media-imdb.com/suggestion/{prefix}/{urllib.parse.quote(q)}.json"
    try:
        r = get_html(url, headers={"Accept": "application/json"})
        data = r.json()
    except Exception:
        return []
    out = []
    for item in data.get("d", []):
        imdb_id = item.get("id")
        t = item.get("l")
        y = item.get("y")
        kind = item.get("qid") or item.get("q") or ""
        if not imdb_id or not imdb_id.startswith("tt"):
            continue
        out.append({"tt": imdb_id, "title": t or "", "year": int(y) if y else None, "kind": str(kind).lower()})
    return out

def candidate_title_similarity(c_title: str, w_title: str) -> float:
    a = normalize_title(c_title); b = normalize_title(w_title)
    if not a or not b:
        return 0.0
    ratio = SequenceMatcher(None, a, b).ratio() * 100.0
    sa, sb = set(a.split()), set(b.split())
    jacc = (len(sa & sb) / max(1, len(sa | sb))) * 100.0 if (sa or sb) else 0.0
    return 0.6 * ratio + 0.4 * jacc

def find_best_imdb(title_clean: str, year_hint: str | None):
    candidates = imdb_suggest(title_clean)
    if not candidates:
        return None, None, None

    def kind_penalty(kind: str) -> int:
        k = (kind or "").lower()
        if "videogame" in k or "video game" in k: return -25
        if "tv" in k or "series" in k or "episode" in k or "mini" in k: return -15
        return 0

    want_year = int(year_hint) if year_hint and year_hint.isdigit() else None
    scored = []
    for c in candidates:
        sim = candidate_title_similarity(c["title"], title_clean)
        yp = 0
        if want_year is not None:
            if c["year"] is None:
                yp = -10
            else:
                dy = abs(c["year"] - want_year)
                yp = {0: 30, 1: 10}.get(dy, -20)
        kp = kind_penalty(c.get("kind", ""))
        scored.append((sim + yp + kp, sim, c))

    if want_year is not None:
        exact = [(s, sim, c) for (s, sim, c) in scored if c["year"] == want_year and sim >= 60]
        if exact:
            _, _, c = max(exact, key=lambda x: x[0])
            return c["tt"], c["title"], str(c["year"]) if c["year"] else None
        near = [(s, sim, c) for (s, sim, c) in scored if (c["year"] is not None and abs(c["year"] - want_year) == 1 and sim >= 80)]
        if near:
            _, _, c = max(near, key=lambda x: x[0])
            return c["tt"], c["title"], str(c["year"]) if c["year"] else None
        return None, None, None

    good = [(s, sim, c) for (s, sim, c) in scored if sim >= 80]
    if good:
        _, _, c = max(good, key=lambda x: x[0])
        return c["tt"], c["title"], str(c["year"]) if c["year"] else None
    return None, None, None

def fetch_title_year_by_tt(tt):
    url = f"https://www.imdb.com/title/{tt}/"
    try:
        html = get_html(url, headers={"Accept": "text/html"}).text
        m = re.search(r"<title>\s*(.*?)\s*-\s*IMDb\s*</title>", html, re.IGNORECASE | re.DOTALL)
        if m:
            t = re.sub(r"\s+", " ", m.group(1)).strip()
            m2 = re.search(r"^(.*)\((\d{4})\)\s*$", t)
            if m2: return m2.group(1).strip(), m2.group(2)
            return t, None
        m = re.search(r'property=["\']og:title["\']\s+content=["\'](.*?)["\']', html, re.IGNORECASE)
        if m:
            t = re.sub(r"\s+", " ", m.group(1)).strip()
            m2 = re.search(r"^(.*)\((\d{4})\)\s*", t)
            if m2: return m2.group(1).strip(), m2.group(2)
            return t, None
    except Exception:
        pass
    return None, None

# ----------------------------- YIFY page helpers -----------------------------
def fetch_movie_page_any(preferred_base, tt):
    bases = [preferred_base] + [b for b in MIRRORS if b != preferred_base]
    for base in bases:
        url = urllib.parse.urljoin(base, f"/movie-imdb/{tt}")
        try:
            html = get_html(url).text
            return html, base
        except Exception:
            continue
    raise RuntimeError(f"All mirrors failed for IMDb {tt}")

def find_lang_slug(movie_html, lang_token):
    pattern = rf'/subtitles/[^"\s<>]+-(?:{lang_token})-(?:yify|yts)-\d+'
    m = re.search(pattern, movie_html, flags=re.IGNORECASE)
    return m.group(0) if m else None

def find_zip_link(sub_html, base):
    m = re.search(r'(https?://[^"\s)]+/subtitle/[^"\s)]+\.zip)', sub_html, flags=re.IGNORECASE)
    if m: return m.group(1)
    m = re.search(r'(/subtitle/[^"\s)]+\.zip)', sub_html, flags=re.IGNORECASE)
    if m: return urllib.parse.urljoin(base, m.group(1))
    return None

def download_zip(url, dest_dir):
    r = get_html(url, stream=True)
    fname = os.path.basename(urllib.parse.urlparse(url).path) or "subtitle.zip"
    if not fname.lower().endswith(".zip"):
        fname += ".zip"
    tmp = os.path.join(dest_dir, f"__dl_{fname}")
    with open(tmp, "wb") as f:
        for chunk in r.iter_content(chunk_size=8192):
            if chunk:
                f.write(chunk)
    return tmp

def same_drive(p1, p2):
    d1 = os.path.splitdrive(os.path.abspath(p1))[0].lower()
    d2 = os.path.splitdrive(os.path.abspath(p2))[0].lower()
    return d1 == d2

def safe_move(src_path, dest_path, overwrite=False):
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    target = dest_path

    if not same_drive(src_path, dest_path):
        if overwrite and os.path.exists(dest_path):
            try: os.remove(dest_path)
            except OSError: pass
        elif not overwrite and os.path.exists(dest_path):
            root, ext = os.path.splitext(dest_path)
            i = 1
            while True:
                cand = f"{root} ({i}){ext}"
                if not os.path.exists(cand):
                    target = cand; break
                i += 1
        shutil.copy2(src_path, target)
        try: os.remove(src_path)
        except OSError: pass
        return target

    if overwrite:
        try: os.replace(src_path, dest_path)
        except OSError:
            if os.path.exists(dest_path):
                try: os.remove(dest_path)
                except OSError: pass
            shutil.move(src_path, dest_path)
        return dest_path

    if os.path.exists(dest_path):
        root, ext = os.path.splitext(dest_path)
        i = 1
        while True:
            cand = f"{root} ({i}){ext}"
            if not os.path.exists(cand):
                target = cand; break
            i += 1
    os.replace(src_path, target)
    return target

def extract_and_rename(zip_path, movie_filepath, lang_suffix, overwrite=False):
    base, _ = os.path.splitext(movie_filepath)
    target = os.path.abspath(f"{base}.{lang_suffix}.srt")
    tempdir = tempfile.mkdtemp(prefix="subs_")
    found = []
    try:
        with ZipFile(zip_path) as zf:
            for member in zf.namelist():
                if member.lower().endswith(".srt") and not member.endswith("/"):
                    zf.extract(member, tempdir)
                    found.append(os.path.join(tempdir, *member.split("/")))
        if not found:
            return False, "ZIP had no .srt"
        found.sort(key=lambda p: os.path.getsize(p), reverse=True)
        chosen = found[0]
        saved = safe_move(chosen, target, overwrite=overwrite)
        return True, os.path.basename(saved)
    except Exception as e:
        return False, str(e)
    finally:
        try: shutil.rmtree(tempdir, ignore_errors=True)
        except: pass
        try: os.remove(zip_path)
        except: pass

# ----------------------------- Persistence helpers -----------------------------
def load_mapping(folder: str) -> dict:
    path = os.path.join(folder, MAP_FILENAME)
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data if isinstance(data, dict) else {}
    except Exception:
        return {}

def save_mapping(folder: str, mapping: dict):
    try:
        with open(os.path.join(folder, MAP_FILENAME), "w", encoding="utf-8") as f:
            json.dump(mapping, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def load_last_folder(default_path: str) -> str:
    try:
        with open(APP_STATE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            p = data.get("last_folder")
            if p and os.path.isdir(p):
                return p
    except Exception:
        pass
    return default_path

def save_last_folder(path: str):
    try:
        with open(APP_STATE_FILE, "w", encoding="utf-8") as f:
            json.dump({"last_folder": path}, f)
    except Exception:
        pass

# ----------------------------- UI utilities -----------------------------
def make_checkbox_images(size: int = 16):
    unchecked = tk.PhotoImage(width=size, height=size)
    unchecked.put("white", to=(0, 0, size, size))
    border = "gray20"
    unchecked.put(border, to=(0, 0, size, 1))
    unchecked.put(border, to=(0, size - 1, size, size))
    unchecked.put(border, to=(0, 0, 1, size))
    unchecked.put(border, to=(size - 1, 0, size, size))

    checked = tk.PhotoImage(width=size, height=size)
    checked.tk.call(checked, "copy", unchecked)
    mark = "#2e7d32"
    pts = [(3, size // 2), (4, size // 2 + 1), (5, size // 2 + 2),
           (6, size // 2 + 3), (7, size // 2 + 2),
           (8, size // 2 + 1), (9, size // 2), (10, size // 2 - 1)]
    for (x, y) in pts:
        checked.put(mark, to=(x, y, x + 1, y + 1))
        checked.put(mark, to=(x, y - 1, x + 1, y))
    return unchecked, checked

def make_donate_image(width=130, height=38):
    radius = height // 2
    top = (255, 187, 71); mid = (247, 162, 28); bot = (225, 140, 22)
    im = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    dr = ImageDraw.Draw(im)
    for y in range(height):
        if y < height * 0.6:
            t = y / (height * 0.6)
            col = tuple(int(top[i] * (1 - t) + mid[i] * t) for i in range(3)) + (255,)
        else:
            t = (y - height * 0.6) / (height * 0.4)
            col = tuple(int(mid[i] * (1 - t) + bot[i] * t) for i in range(3)) + (255,)
        dr.line([(0, y), (width, y)], fill=col)
    mask = Image.new("L", (width, height), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, width - 1, height - 1], radius=radius, fill=255)
    im.putalpha(mask)
    highlight = Image.new("RGBA", (width, height), (255, 255, 255, 0))
    ImageDraw.Draw(highlight).rounded_rectangle([2, 2, width - 3, height // 2], radius=radius - 2, fill=(255, 255, 255, 70))
    im = Image.alpha_composite(im, highlight)
    ImageDraw.Draw(im).rounded_rectangle([0.5, 0.5, width - 1.5, height - 1.5], radius=radius, outline=(200, 120, 20, 255), width=2)
    bio = BytesIO(); im.save(bio, format="PNG"); bio.seek(0)
    return tk.PhotoImage(data=bio.read())

class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip = None
        widget.bind("<Enter>", self.show); widget.bind("<Leave>", self.hide)
    def show(self, _=None):
        if self.tip or not self.text: return
        x = self.widget.winfo_rootx() + 25
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 10
        self.tip = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True); tw.wm_geometry(f"+{x}+{y}")
        tk.Label(tw, text=self.text, background="#ffffe0", relief="solid",
                 borderwidth=1, font=("Segoe UI", 9)).pack(ipadx=6, ipady=2)
    def hide(self, _=None):
        if self.tip: self.tip.destroy(); self.tip = None

# ----------------------------- GUI -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("YIFY Subtitle Downloader v1.1")
        self.geometry("1220x800"); self.minsize(1060, 700)

        # Window icon (icon.ico if present)
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            ico_path = os.path.join(script_dir, "icon.ico")
            if os.path.exists(ico_path):
                self.iconbitmap(default=ico_path)  # Windows .ico
            else:
                png_path = os.path.join(script_dir, "icon.png")
                if os.path.exists(png_path):
                    self._icon_png = tk.PhotoImage(file=png_path)
                    self.iconphoto(True, self._icon_png)
        except Exception:
            pass

        # load last folder (default = cwd)
        last_folder = load_last_folder(os.getcwd())
        self.folder_var = tk.StringVar(value=last_folder)
        self.folder_var.trace_add("write", lambda *_: save_last_folder(self.folder_var.get()))

        self.lang_en_var = tk.BooleanVar(value=True)
        self.lang_ar_var = tk.BooleanVar(value=True)
        self.overwrite_var = tk.BooleanVar(value=False)

        # threads/cancel
        self.match_thread = None; self.download_thread = None
        self.match_cancel = threading.Event(); self.download_cancel = threading.Event()

        # selection state (checkbox images)
        self.img_unchecked, self.img_checked = make_checkbox_images(16)
        self.checked = set()  # iids

        # data
        self.rows = {}         # iid -> row dict
        self.mapping = {}      # per-folder mapping

        self.build_ui()
        self.center_window()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---- close/save ----
    def on_close(self):
        save_last_folder(self.folder_var.get())
        self.destroy()

    # ---- centering ----
    def center_window(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    def center_toplevel(self, win):
        win.update_idletasks()
        w, h = win.winfo_width(), win.winfo_height()
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        win.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    # ---- progress helpers ----
    def progress_start(self, phase: str, total: int):
        self.pb_phase = phase; self.pb_total = max(0, int(total)); self.pb_value = 0
        self.pb.configure(mode="determinate", maximum=max(self.pb_total, 1), value=0)
        self.pb_label.configure(text=f"{phase}: 0/{self.pb_total}")
        self.update_idletasks()

    def progress_step(self, inc: int = 1):
        if getattr(self, "pb_total", 0) <= 0: return
        self.pb_value = min(self.pb_total, getattr(self, "pb_value", 0) + inc)
        self.pb.configure(value=self.pb_value)
        self.pb_label.configure(text=f"{self.pb_phase}: {self.pb_value}/{self.pb_total}")
        self.update_idletasks()

    def progress_finish(self, canceled=False):
        if getattr(self, "pb_total", 0) > 0:
            self.pb.configure(value=self.pb_total)
            self.pb_label.configure(text=f"{self.pb_phase}: {self.pb_total}/{self.pb_total}" + (" (canceled)" if canceled else ""))
        else:
            self.pb_label.configure(text="Idle")
        self.update_idletasks()

    # ---- UI ----
    def build_ui(self):
        # Top bar
        top = ttk.Frame(self); top.pack(fill="x", padx=10, pady=10)
        ttk.Label(top, text="Movies Folder:").pack(side="left")
        self.folder_entry = ttk.Entry(top, textvariable=self.folder_var, width=80); self.folder_entry.pack(side="left", padx=6)
        ttk.Button(top, text="Browse…", command=self.on_browse).pack(side="left", padx=4)
        ttk.Button(top, text="List Files", command=self.on_list_files).pack(side="left", padx=8)
        self.btn_fetch = ttk.Button(top, text="Fetch IMDb ID", command=self.on_fetch_clicked); self.btn_fetch.pack(side="left", padx=12)
        ttk.Button(top, text="Include: All", command=self.select_all).pack(side="left", padx=6)
        ttk.Button(top, text="Include: None", command=self.select_none).pack(side="left", padx=4)

        # Tree + scrollbars
        wrap = ttk.Frame(self); wrap.pack(fill="both", expand=True, padx=10, pady=(6, 4))
        cols = ("filename", "guess", "tt", "movietitle", "year", "status")
        self.tree = ttk.Treeview(wrap, columns=cols, show="tree headings", selectmode="browse")

        self.tree.heading("#0", text="Include", anchor="center")
        self.tree.heading("filename", text="Filename", anchor="w")
        self.tree.heading("guess", text="Guessed Title (editable)", anchor="w")
        self.tree.heading("tt", text="IMDb ID (editable)", anchor="center")
        self.tree.heading("movietitle", text="Matched Title", anchor="w")
        self.tree.heading("year", text="Year", anchor="center")
        self.tree.heading("status", text="Status", anchor="w")

        font = tkfont.nametofont("TkDefaultFont")
        include_w = font.measure("Include") + 20
        self.tree.column("#0", width=include_w, minwidth=include_w, anchor="center", stretch=False)  # fixed
        self.tree.column("filename", width=340, minwidth=180, anchor="w", stretch=False)
        self.tree.column("guess", width=280, minwidth=160, anchor="w", stretch=False)
        self.tree.column("tt", width=140, minwidth=90, anchor="center", stretch=False)
        self.tree.column("movietitle", width=280, minwidth=160, anchor="w", stretch=False)
        self.tree.column("year", width=70, minwidth=50, anchor="center", stretch=False)
        self.tree.column("status", width=250, minwidth=160, anchor="w", stretch=False)
        self.tree["displaycolumns"] = cols

        ysb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        xsb = ttk.Scrollbar(wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscroll=ysb.set, xscroll=xsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew"); ysb.grid(row=0, column=1, sticky="ns"); xsb.grid(row=1, column=0, sticky="ew")
        wrap.rowconfigure(0, weight=1); wrap.columnconfigure(0, weight=1)

        # Mouse/column handlers
        self._block_resize_select = False
        self.tree.bind("<Button-1>", self._on_mouse_down, add="+")
        self.tree.bind("<B1-Motion>", self._on_mouse_drag, add="+")
        self.tree.bind("<ButtonRelease-1>", self._on_mouse_up, add="+")
        self.tree.bind("<Double-Button-1>", self._on_double_click_header, add="+")  # Excel auto-fit
        self.tree.bind("<Double-1>", self.on_tree_double_click)  # inline edit
        self.tree.bind("<Button-3>", self.on_right_click)

        # Context menu
        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Edit", command=self.ctx_edit)
        self.menu.add_command(label="Find IMDb ID", command=self.ctx_find_imdb)
        self.menu.add_command(label="Download Subtitle", command=self.ctx_download_one)
        self.menu_row = None

        # Bottom: options + Download + signature + donate
        bottom = ttk.Frame(self); bottom.pack(fill="x", padx=10, pady=(2, 6))
        left = ttk.Frame(bottom); left.pack(side="left")
        self.chk_en = ttk.Checkbutton(left, text="English (.en)", variable=self.lang_en_var)
        self.chk_ar = ttk.Checkbutton(left, text="Arabic (.ar)", variable=self.lang_ar_var)
        self.chk_over = ttk.Checkbutton(left, text="Overwrite existing subtitles", variable=self.overwrite_var)
        self.chk_en.pack(side="left", padx=(0, 8)); self.chk_ar.pack(side="left", padx=(0, 12)); self.chk_over.pack(side="left", padx=(0, 12))
        self.btn_download = ttk.Button(bottom, text="Download Subtitles", command=self.on_download_clicked)
        self.btn_download.pack(side="left", padx=(10, 0))

        right = ttk.Frame(bottom); right.pack(side="right")
        # Donate pill with centered text
        self.donate_img = make_donate_image(130, 38)
        self.btn_donate = tk.Button(
            right, image=self.donate_img, text="Donate", compound="center",
            font=("Segoe UI", 11, "bold"), fg="#0f3462", activeforeground="#0f3462",
            bd=0, highlightthickness=0, cursor="hand2", relief="flat",
            command=lambda: self._open_donate_link()
        )
        self.btn_donate.pack(side="right", padx=(8, 0))
        ToolTip(self.btn_donate, "Support development with a small donation")

        # Signature text (optional icon.png)
        try:
            icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.png")
            if os.path.exists(icon_path):
                self.icon_img = tk.PhotoImage(file=icon_path)
                ttk.Label(right, image=self.icon_img).pack(side="right", padx=(8, 6))
        except Exception:
            self.icon_img = None
        ttk.Label(right, text="made by Boyaqoub - ilukezippo@gmail.com",
                  font=("Segoe UI", 10, "bold")).pack(side="right")

        # Progress + log
        pwrap = ttk.Frame(self); pwrap.pack(fill="x", padx=10, pady=(0, 4))
        self.pb_label = ttk.Label(pwrap, text="Idle"); self.pb_label.pack(side="left")
        self.pb = ttk.Progressbar(pwrap, orient="horizontal", mode="determinate"); self.pb.pack(fill="x", expand=True, padx=10)

        logf = ttk.LabelFrame(self, text="Log"); logf.pack(fill="both", expand=False, padx=10, pady=(0, 10))
        self.log_text = tk.Text(logf, height=8, wrap="none"); self.log_text.pack(fill="both", expand=True)

        self.log("Tip: List files, edit Title/IMDb ID if needed, click 'Fetch IMDb ID', then 'Download Subtitles'.")

    # ---- Mouse/columns ----
    def _on_mouse_down(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region == "separator":
            col_left = self.tree.identify_column(event.x - 1)
            if col_left == "#0":
                self._block_resize_select = True
                return "break"
            self._block_resize_select = False
            return
        if region == "tree":
            iid = self.tree.identify_row(event.y)
            col = self.tree.identify_column(event.x)
            if iid and col == "#0":
                if iid in self.checked:
                    self.checked.remove(iid); self.tree.item(iid, image=self.img_unchecked)
                else:
                    self.checked.add(iid); self.tree.item(iid, image=self.img_checked)
                return "break"
        self._block_resize_select = False

    def _on_mouse_drag(self, event):
        if self._block_resize_select:
            return "break"

    def _on_mouse_up(self, _):
        self._block_resize_select = False

    def _on_double_click_header(self, event):
        """Excel-like: double-click header right edge to auto-fit that column (except Include/#0)."""
        region = self.tree.identify("region", event.x, event.y)

        # If we double-clicked directly on a separator, identify the column to the left
        if region == "separator":
            col_left = self.tree.identify_column(event.x - 1)
            if col_left and col_left != "#0":
                self.autofit_column(col_left)
            return "break"

        # Some Tk builds report 'heading' (not 'separator') on a double-click near the edge.
        # Detect if the click was within a few pixels from the right edge of that heading.
        if region == "heading":
            clicked_col = self.tree.identify_column(event.x)
            if not clicked_col or clicked_col == "#0":
                return  # do not auto-fit the Include column

            # Compute right edge of that header in widget coords, accounting for horizontal scroll
            order = ["#0"] + list(self.tree["displaycolumns"])
            total_w = sum(int(self.tree.column(cid, "width")) for cid in order)
            # xview returns (start_frac, end_frac) of the *total* width currently scrolled
            scroll_px = int(self.tree.xview()[0] * total_w)

            xleft = -scroll_px
            right_edge = None
            for cid in order:
                w = int(self.tree.column(cid, "width"))
                xright = xleft + w
                if cid == clicked_col:
                    right_edge = xright
                    break
                xleft = xright

            if right_edge is not None and abs(event.x - right_edge) <= 6:
                self.autofit_column(clicked_col)
                return "break"

    def autofit_column(self, col_id: str):
        heading = self.tree.heading(col_id, "text") or ""
        font = tkfont.nametofont("TkDefaultFont")
        pad = 24  # breathing room
        max_px = font.measure(heading)
        for item in self.tree.get_children(""):
            val = self.tree.set(item, col_id) or ""
            px = font.measure(val)
            if px > max_px:
                max_px = px
        minw = int(self.tree.column(col_id, "minwidth") or 20)
        self.tree.column(col_id, width=max(minw, max_px + pad))

    # ---- Logging ----
    def log(self, msg: str):
        self.log_text.insert("end", msg + "\n"); self.log_text.see("end"); self.update_idletasks()

    def set_status(self, iid, text):
        self.tree.set(iid, "status", text); self.update_idletasks()

    # ---- File ops ----
    def on_browse(self):
        start = self.folder_var.get().strip() or os.getcwd()
        folder = filedialog.askdirectory(initialdir=start)
        if folder:
            self.folder_var.set(folder)  # trace saves it

    def on_list_files(self):
        folder = self.folder_var.get().strip() or os.getcwd()
        if not os.path.isdir(folder):
            messagebox.showerror("Error", "Invalid folder."); return
        self.mapping = load_mapping(folder)
        for iid in self.tree.get_children(): self.tree.delete(iid)
        self.rows.clear(); self.checked.clear()

        files = sorted(os.listdir(folder))
        cnt = 0
        for fname in files:
            path = os.path.join(folder, fname)
            if os.path.isfile(path) and os.path.splitext(fname)[1].lower() in VIDEO_EXTS:
                guess, _ = filename_to_title_and_year(fname)
                tt = title = year = ""
                status = "Ready"
                if fname in self.mapping:
                    m = self.mapping[fname] or {}
                    tt = m.get("tt", "") or ""; title = m.get("title", "") or ""; year = m.get("year", "") or ""
                    if tt: status = "Matched (saved)"
                iid = self.tree.insert("", "end", text="", image=self.img_checked,
                                       values=(fname, guess, tt, title, year, status))
                self.checked.add(iid)
                self.rows[iid] = {"path": path, "guess": guess, "tt": tt, "title": title, "year": year, "base": ""}
                cnt += 1
        self.log(f"Found {cnt} video files.")

    # ---- Include all/none ----
    def select_all(self):
        for iid in self.tree.get_children(""):
            self.checked.add(iid); self.tree.item(iid, image=self.img_checked)

    def select_none(self):
        for iid in self.tree.get_children(""):
            self.checked.discard(iid); self.tree.item(iid, image=self.img_unchecked)

    # ---- Inline edit ----
    def on_tree_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell": return
        rowid = self.tree.identify_row(event.y); colid = self.tree.identify_column(event.x)
        if not rowid or not colid: return
        colname = self.tree["columns"][int(colid[1:]) - 1]
        if colname not in ("guess", "tt"): return

        x, y, w, h = self.tree.bbox(rowid, colname)
        value = self.tree.set(rowid, colname)
        entry = ttk.Entry(self.tree); entry.insert(0, value); entry.select_range(0, "end"); entry.focus_set()
        entry.place(x=x, y=y, width=w, height=h)

        def finish_edit(_=None):
            new_val = entry.get().strip(); entry.destroy()
            self.tree.set(rowid, colname, new_val)
            if colname == "guess":
                self.rows[rowid]["guess"] = new_val
                self.rows[rowid]["title"] = ""; self.rows[rowid]["year"] = ""; self.rows[rowid]["base"] = ""
                self.tree.set(rowid, "movietitle", ""); self.tree.set(rowid, "year", "")
            else:
                self.rows[rowid]["tt"] = new_val
                threading.Thread(target=self._refresh_title_from_tt, args=(rowid, new_val), daemon=True).start()

        entry.bind("<Return>", finish_edit)
        entry.bind("<Escape>", lambda e: (entry.destroy()))
        entry.focus_force()

    # ---- Context menu ----
    def on_right_click(self, event):
        rowid = self.tree.identify_row(event.y)
        if not rowid: return
        self.menu_row = rowid
        try: self.menu.tk_popup(event.x_root, event.y_root)
        finally: self.menu.grab_release()

    def ctx_edit(self):
        rid = self.menu_row
        if not rid: return
        row = self.rows[rid]
        self._open_edit_dialog(rid, row)

    def ctx_find_imdb(self):
        rid = self.menu_row
        if not rid: return
        row = self.rows[rid]
        if row.get("tt"):
            self.set_status(rid, "IMDb set (skipped)"); self.log(f"[SKIP] {os.path.basename(row['path'])}: already has ID.")
            return
        threading.Thread(target=self._find_match_for_row, args=(rid, row), daemon=True).start()

    def ctx_download_one(self):
        rid = self.menu_row
        if not rid: return
        row = self.rows[rid]
        langs = []; 
        if self.lang_en_var.get(): langs.append(("english", "en"))
        if self.lang_ar_var.get(): langs.append(("arabic", "ar"))
        if not langs:
            messagebox.showinfo("Info", "Please check at least one language (English and/or Arabic)."); return
        self.progress_start("Downloading", len(langs))
        threading.Thread(target=self._download_for_rows, args=([(rid, row)], langs, False), daemon=True).start()

    def _open_edit_dialog(self, rid, row):
        win = tk.Toplevel(self); win.title("Edit row"); win.transient(self); win.grab_set()
        ttk.Label(win, text="Guessed Title:").grid(row=0, column=0, sticky="e", padx=6, pady=6)
        e_title = ttk.Entry(win, width=50); e_title.insert(0, row["guess"]); e_title.grid(row=0, column=1, padx=6, pady=6)
        ttk.Label(win, text="IMDb ID:").grid(row=1, column=0, sticky="e", padx=6, pady=6)
        e_tt = ttk.Entry(win, width=30); e_tt.insert(0, row.get("tt","")); e_tt.grid(row=1, column=1, padx=6, pady=6, sticky="w")

        def save():
            new_title = e_title.get().strip(); new_tt = e_tt.get().strip()
            self.rows[rid]["guess"] = new_title; self.rows[rid]["tt"] = new_tt
            self.tree.set(rid, "guess", new_title); self.tree.set(rid, "tt", new_tt)
            if new_tt: threading.Thread(target=self._refresh_title_from_tt, args=(rid, new_tt), daemon=True).start()
            win.destroy()

        btns = ttk.Frame(win); btns.grid(row=2, column=0, columnspan=2, pady=8)
        ttk.Button(btns, text="Save", command=save).pack(side="left", padx=6)
        ttk.Button(btns, text="Cancel", command=win.destroy).pack(side="left", padx=6)
        win.bind("<Return>", lambda e: save()); self.center_toplevel(win)

    # ---- Fetch IMDb (toggleable) ----
    def on_fetch_clicked(self):
        if self.match_thread and self.match_thread.is_alive():
            self.match_cancel.set(); self.log("Cancel requested: Fetch IMDb ID"); return

        work = [(iid, row) for iid, row in self.rows.items() if (iid in self.checked) and not row.get("tt")]
        if not work:
            messagebox.showinfo("Info", "No rows selected (or all selected rows already have IMDb IDs)."); return

        self.match_cancel.clear(); self.btn_fetch.config(text="Cancel")
        self.log("Finding IMDb matches (strict title + year)…")
        self.progress_start("Matching", len(work))
        self.match_thread = threading.Thread(target=self._thread_find_matches, args=(work,), daemon=True); self.match_thread.start()

    def _thread_find_matches(self, items):
        canceled = False
        for iid, row in items:
            if self.match_cancel.is_set(): canceled = True; break
            self._find_match_for_row(iid, row); self.progress_step(1)
        self.progress_finish(canceled=canceled)
        self.after(0, lambda: self.btn_fetch.config(text="Fetch IMDb ID"))
        self.log("Matching canceled." if canceled else "Matching complete.")

    def _save_row_mapping(self, iid):
        row = self.rows[iid]; fname = os.path.basename(row["path"])
        if not row.get("tt"): return
        folder = os.path.dirname(row["path"])
        cur = load_mapping(folder)
        cur[fname] = {"tt": row.get("tt",""), "title": row.get("title",""), "year": row.get("year","")}
        save_mapping(folder, cur)

    def _find_match_for_row(self, iid, row):
        fname = os.path.basename(row["path"])
        title_guess, year_hint = filename_to_title_and_year(fname)
        if row.get("guess"): title_guess = row["guess"].strip() or title_guess

        self.set_status(iid, "Searching IMDb…")
        try:
            tt, mtitle, myear = find_best_imdb(title_guess, year_hint)
            if not tt:
                self.set_status(iid, "Not found"); self.log(f"[NO MATCH] {fname} (title='{title_guess}', year={year_hint or 'N/A'})"); return
            row["tt"] = tt; row["title"] = mtitle or ""; row["year"] = myear or (year_hint or ""); row["base"] = MIRRORS[0]
            self.tree.set(iid, "tt", tt); self.tree.set(iid, "movietitle", row["title"]); self.tree.set(iid, "year", row["year"])
            self.set_status(iid, "Matched"); self._save_row_mapping(iid)
            self.log(f"[MATCH] {fname} → {row['title']} ({row['year']}) [{tt}]")
        except Exception as e:
            self.set_status(iid, f"Error: {e}"); self.log(f"[ERROR] Finding match for {fname}: {e}")

    def _refresh_title_from_tt(self, iid, tt):
        if not tt: return
        self.set_status(iid, "Fetching title by IMDb ID…")
        title, year = fetch_title_year_by_tt(tt)
        if title:
            self.rows[iid]["title"] = title
            if year: self.rows[iid]["year"] = year
            self.tree.set(iid, "movietitle", title); 
            if year: self.tree.set(iid, "year", year)
            self.set_status(iid, "Matched (manual)"); self._save_row_mapping(iid)
            self.log(f"[ID→TITLE] {os.path.basename(self.rows[iid]['path'])} → {title} ({year or '?'}) [{tt}]")
        else:
            self.set_status(iid, "ID fetch failed"); self.log(f"[ERROR] Could not fetch title for {tt}")

    # ---- Download (toggleable) ----
    def on_download_clicked(self):
        if self.download_thread and self.download_thread.is_alive():
            self.download_cancel.set(); self.log("Cancel requested: Download Subtitles"); return
        if not (self.lang_en_var.get() or self.lang_ar_var.get()):
            messagebox.showinfo("Info", "Please check at least one language (English and/or Arabic)."); return
        work = [(iid, row) for iid, row in self.rows.items() if (iid in self.checked) and row.get("tt")]
        if not work:
            messagebox.showinfo("Info", "No matched rows selected (IMDb ID missing)."); return
        langs = []; 
        if self.lang_en_var.get(): langs.append(("english", "en"))
        if self.lang_ar_var.get(): langs.append(("arabic", "ar"))
        self.download_cancel.clear(); self.btn_download.config(text="Cancel")
        self.log("Downloading subtitles…"); self.progress_start("Downloading", len(work) * len(langs))
        self.download_thread = threading.Thread(target=self._download_for_rows, args=(work, langs, True), daemon=True)
        self.download_thread.start()

    def _subtitle_exists(self, movie_path, suffix):
        base, _ = os.path.splitext(movie_path)
        return os.path.exists(f"{base}.{suffix}.srt")

    def _download_for_rows(self, items, langs, toggle_button_after):
        overwrite = self.overwrite_var.get(); canceled = False
        for iid, row in items:
            if self.download_cancel.is_set(): canceled = True; break
            fname = os.path.basename(row["path"]); tt = row["tt"]
            base_pref = row.get("base") or MIRRORS[0]
            self.set_status(iid, "Opening movie page…")
            try:
                movie_html, base = fetch_movie_page_any(base_pref, tt)
            except Exception as e:
                self.set_status(iid, "Movie page error"); self.log(f"[ERROR] Movie page for {fname}: {e}")
                for _ in langs: self.progress_step(1); continue

            any_saved = False
            for lang_token, suffix in langs:
                if self.download_cancel.is_set(): canceled = True; break
                if not overwrite and self._subtitle_exists(row["path"], suffix):
                    self.log(f"[SKIP] {fname}: {suffix}.srt already exists."); self.progress_step(1); continue
                self.set_status(iid, f"{lang_token.title()}…")
                slug = find_lang_slug(movie_html, lang_token)
                if not slug: self.log(f"[MISS] {fname}: no {lang_token} slug."); self.progress_step(1); continue

                sub_html = None; sub_base = base
                for b in [base] + [m for m in MIRRORS if m != base]:
                    try: sub_html = get_html(urllib.parse.urljoin(b, slug)).text; sub_base = b; break
                    except Exception: continue
                if not sub_html: self.log(f"[ERROR] {fname}: cannot open {lang_token} subtitle page."); self.progress_step(1); continue

                zip_url = find_zip_link(sub_html, sub_base)
                if not zip_url: self.log(f"[ERROR] {fname}: no .zip link for {lang_token}."); self.progress_step(1); continue
                try:
                    tmp_zip = download_zip(zip_url, dest_dir=os.path.dirname(row["path"]))
                    ok, result = extract_and_rename(tmp_zip, row["path"], suffix, overwrite=overwrite)
                    if ok: self.log(f"[OK] {fname}: saved {result}"); any_saved = True
                    else:  self.log(f"[ERROR] {fname}: {lang_token} – {result}")
                except Exception as e:
                    self.log(f"[ERROR] {fname}: {lang_token} – {e}")
                finally:
                    self.progress_step(1)
            self.set_status(iid, "Done" if any_saved else "No subs")

        self.progress_finish(canceled=canceled)
        if toggle_button_after:
            self.after(0, lambda: self.btn_download.config(text="Download Subtitles"))
        self.log("Download canceled." if canceled else "Download complete.")

    # ---- Donation link ----
    def _open_donate_link(self):
        import webbrowser
        webbrowser.open("https://buymeacoffee.com/ilukezippo")

# ----------------------------- main -----------------------------
if __name__ == "__main__":
    app = App()
    app.mainloop()
