# YIFY-Subtitle-Downloader.py
# (Only the top/additions and the bottom credit icon loader changed)

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

import requests
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ---------- NEW: resource path helper for PyInstaller ----------
import sys
def resource_path(relative_path: str) -> str:
    """Get absolute path to resource (works in dev and in PyInstaller bundle)."""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# ----------------------------- Config -----------------------------
MIRRORS = [
    "https://yifysubtitles.ch",
    "https://yts-subs.com",
]

VIDEO_EXTS = {".mp4", ".mkv", ".avi", ".mov", ".wmv", ".flv", ".m4v"}
LANG_TOKENS = {"english": "en", "arabic": "ar"}  # site language -> filename suffix
MAP_FILENAME = ".yify_imdb_map.json"            # stored in the selected folder

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

CHECKED = "☑"
UNCHECKED = "☐"

# …………… (unchanged app code) ……………

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("YIFY Subtitle Downloader v1.0 (IMDb match, EN/AR)")
        self.geometry("1220x800")
        self.minsize(1060, 700)

        self.folder_var = tk.StringVar(value=os.getcwd())
        self.lang_en_var = tk.BooleanVar(value=True)
        self.lang_ar_var = tk.BooleanVar(value=True)
        self.overwrite_var = tk.BooleanVar(value=False)

        # Progress state
        self.pb_total = 0
        self.pb_value = 0
        self.pb_phase = ""

        # Cancellation + running threads
        self.match_thread = None
               # ... rest of your original code stays the same ...

    def build_ui(self):
        # … your existing UI setup above …

        # Bottom controls bar (languages + Overwrite + Download + credit)
        bottom = ttk.Frame(self); bottom.pack(fill="x", padx=10, pady=(2, 6))

        # Credit + icon (NOW loads from resource_path so it works inside EXE)
        credit_wrap = ttk.Frame(bottom)
        credit_wrap.pack(side="right")
        self.icon_img = None
        try:
            icon_path = resource_path("icon.png")
            if os.path.exists(icon_path):
                self.icon_img = tk.PhotoImage(file=icon_path)
                ttk.Label(credit_wrap, image=self.icon_img).pack(side="left", padx=(0, 6))
        except Exception:
            pass
        ttk.Label(credit_wrap, text="made by Boyaqoub - ilukezippo@gmail.com").pack(side="left")

        # … rest of your original UI code unchanged …
        # (progress bar, log, etc.)

# ----------------------------- main -----------------------------
if __name__ == "__main__":
    app = App()
    app.mainloop()
