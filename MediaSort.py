"""
MediaSort — sorts images, videos, documents, and audio into organised sub-folders.

┌─────────────────────────────────────────────────────────────────┐
│  IMAGE categories  (checked in priority order)                  │
│   1. Wallpapers      → filename contains "wallhaven"            │
│   2. Screenshots     → filename contains "screenshot"           │
│   3. WebP Images     → extension is .webp                       │
│   4. Square Images   → width/height ratio within 5 % of 1:1    │
│   5. Other Images    → everything else                          │
├─────────────────────────────────────────────────────────────────┤
│  VIDEO categories  (checked in priority order)                  │
│   1. Screen Recordings → filename starts with "screen_recording"│
│   2. 16:9 Landscape    → aspect ratio ≈ 16:9                    │
│   3. 9:16 Portrait     → aspect ratio ≈ 9:16                    │
│   4. Other             → anything else                          │
├─────────────────────────────────────────────────────────────────┤
│  DOCUMENT categories                                            │
│   1. Word Documents  → .doc .docx .odt .rtf                    │
│   2. Spreadsheets    → .xls .xlsx .ods .csv .tsv               │
│   3. Presentations   → .ppt .pptx .odp                         │
│   4. PDFs            → .pdf                                     │
│   5. Text & Markdown → .txt .md .rst .log                      │
│   6. Code & Scripts  → .py .js .ts .html .css .json .xml ...   │
│   7. Archives        → .zip .rar .7z .tar .gz .bz2 .xz         │
│   8. Other Docs      → everything else                          │
├─────────────────────────────────────────────────────────────────┤
│  AUDIO categories                                               │
│   1. Lossless        → .flac .wav .aiff .aif .alac .ape .wv    │
│   2. Compressed      → .mp3 .aac .ogg .opus .m4a .wma .ac3 ... │
│   3. Playlists       → .m3u .m3u8 .pls .xspf .wpl              │
│   4. Other Audio     → everything else                          │
└─────────────────────────────────────────────────────────────────┘

Dependencies:
  tkinter            — GUI (stdlib, bundled with Python on Windows)
  threading          — background scan (stdlib)
  os, pathlib, shutil, math, subprocess, json  (stdlib)
  Pillow             — image dimensions  (pip install Pillow)
  ffprobe (on PATH)  — video dimensions  (or opencv-python as fallback)
  opencv-python      — optional fallback  (pip install opencv-python)
"""

import os
import json
import math
import shutil
import subprocess
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from PIL import Image as _PILImage
    _PIL_AVAILABLE = True
except ImportError:
    _PIL_AVAILABLE = False


# ══════════════════════════════════════════════════════════════════
# Shared palette
# ══════════════════════════════════════════════════════════════════

PALETTE = {
    "bg":        "#0f0f13",
    "panel":     "#17171f",
    "card":      "#1e1e2a",
    "border":    "#2a2a3a",
    "accent":    "#6c63ff",
    "accent2":   "#ff6584",
    "accent3":   "#43e8a0",
    "accent4":   "#ffd166",
    "text":      "#e8e8f0",
    "subtext":   "#888898",
    "hover":     "#252535",
    "sel":       "#2d2d4a",
    # image category colours
    "wall":      "#b39ddb",   # soft purple  — wallpapers
    "shot":      "#80cbc4",   # teal         — screenshots
    "webp":      "#ffb74d",   # amber        — webp
    "square":    "#4dd0e1",   # cyan         — square images
    "img_other": "#ef9a9a",   # rose         — other images
    # video category colours
    "landscape": "#4fc3f7",   # sky blue     — 16:9
    "portrait":  "#f48fb1",   # pink         — 9:16
    "screen":    "#a5d6a7",   # green        — screen recordings
    "vid_other": "#ffcc80",   # peach        — other videos
    # document category colours
    "word":      "#90caf9",   # light blue   — word docs
    "sheet":     "#a5d6a7",   # light green  — spreadsheets
    "slides":    "#ffcc02",   # yellow       — presentations
    "pdf":       "#ef9a9a",   # rose         — PDFs
    "textmd":    "#ce93d8",   # lilac        — text & markdown
    "code":      "#80deea",   # cyan         — code & scripts
    "archive":   "#ffab91",   # orange       — archives
    "doc_other": "#b0bec5",   # blue-grey    — other docs
    # audio category colours
    "lossless":  "#e6ee9c",   # lime         — lossless audio
    "lossy":     "#80cbc4",   # teal         — compressed audio
    "playlist":  "#b39ddb",   # purple       — playlists
    "aud_other": "#f48fb1",   # pink         — other audio
}


# ══════════════════════════════════════════════════════════════════
# IMAGE constants & logic
# ══════════════════════════════════════════════════════════════════

IMAGE_EXTENSIONS = {
    ".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp",
    ".tiff", ".tif", ".heic", ".heif", ".avif", ".jfif",
    ".ico", ".svg", ".raw", ".cr2", ".nef", ".arw",
    ".dng", ".orf", ".rw2", ".psd", ".xcf",
}

ICАТ_WALLPAPER  = "Wallpapers"
ICАТ_SCREENSHOT = "Screenshots"
ICАТ_WEBP       = "WebP Images"
ICАТ_SQUARE     = "Square Images"
ICАТ_OTHER      = "Other Images"

IMAGE_CATEGORIES = [ICАТ_WALLPAPER, ICАТ_SCREENSHOT, ICАТ_WEBP, ICАТ_SQUARE, ICАТ_OTHER]

# 5 % tolerance: shorter side must be >= 95 % of longer side
SQUARE_TOLERANCE = 0.05

IMAGE_CAT_COLORS = {
    ICАТ_WALLPAPER:  PALETTE["wall"],
    ICАТ_SCREENSHOT: PALETTE["shot"],
    ICАТ_WEBP:       PALETTE["webp"],
    ICАТ_SQUARE:     PALETTE["square"],
    ICАТ_OTHER:      PALETTE["img_other"],
}

IMAGE_CAT_ICONS = {
    ICАТ_WALLPAPER:  "🖼",
    ICАТ_SCREENSHOT: "📸",
    ICАТ_WEBP:       "⚡",
    ICАТ_SQUARE:     "⬛",
    ICАТ_OTHER:      "📁",
}


def _get_image_dimensions(filepath: Path):
    """Return (width, height) via Pillow, or None if unavailable/unreadable."""
    if not _PIL_AVAILABLE:
        return None
    try:
        with _PILImage.open(filepath) as img:
            return img.size  # (width, height)
    except Exception:
        return None


def classify_image(filepath: Path) -> str:
    """Classify one image file. Priority order documented in module docstring."""
    name_lower = filepath.stem.lower()
    ext_lower  = filepath.suffix.lower()

    if "wallhaven" in name_lower:
        return ICАТ_WALLPAPER
    if "screenshot" in name_lower:
        return ICАТ_SCREENSHOT
    if ext_lower == ".webp":
        return ICАТ_WEBP

    dims = _get_image_dimensions(filepath)
    if dims is not None:
        w, h = dims
        if w > 0 and h > 0:
            ratio = min(w, h) / max(w, h)
            if ratio >= (1.0 - SQUARE_TOLERANCE):
                return ICАТ_SQUARE

    return ICАТ_OTHER


# ══════════════════════════════════════════════════════════════════
# VIDEO constants & logic
# ══════════════════════════════════════════════════════════════════

VIDEO_EXTENSIONS = {
    ".mp4", ".mkv", ".avi", ".mov", ".wmv", ".flv", ".webm",
    ".m4v", ".ts", ".mts", ".m2ts", ".vob", ".3gp", ".3g2",
    ".mpg", ".mpeg", ".f4v", ".rm", ".rmvb", ".divx", ".ogv",
}

VCAT_SCREEN_REC = "Screen Recordings"
VCAT_LANDSCAPE  = "16:9 Landscape"
VCAT_PORTRAIT   = "9:16 Portrait"
VCAT_OTHER      = "Other"

VIDEO_CATEGORIES = [VCAT_SCREEN_REC, VCAT_LANDSCAPE, VCAT_PORTRAIT, VCAT_OTHER]

RATIO_TOLERANCE = 0.05   # ±5 % when comparing aspect ratios

VIDEO_CAT_COLORS = {
    VCAT_SCREEN_REC: PALETTE["screen"],
    VCAT_LANDSCAPE:  PALETTE["landscape"],
    VCAT_PORTRAIT:   PALETTE["portrait"],
    VCAT_OTHER:      PALETTE["vid_other"],
}

VIDEO_CAT_ICONS = {
    VCAT_SCREEN_REC: "🎬",
    VCAT_LANDSCAPE:  "🖥",
    VCAT_PORTRAIT:   "📱",
    VCAT_OTHER:      "📂",
}


def _get_dims_ffprobe(path: str):
    """Return (width, height) via ffprobe, or None on failure."""
    cmd = ["ffprobe", "-v", "quiet", "-print_format", "json", "-show_streams", path]
    try:
        result = subprocess.run(cmd, capture_output=True, timeout=10)
        data = json.loads(result.stdout)
        for stream in data.get("streams", []):
            w = stream.get("width")
            h = stream.get("height")
            if w and h:
                return int(w), int(h)
    except Exception:
        pass
    return None


def _get_dims_cv2(path: str):
    """Return (width, height) via OpenCV, or None on failure."""
    try:
        import cv2
        cap = cv2.VideoCapture(path)
        if cap.isOpened():
            w = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
            h = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
            cap.release()
            if w and h:
                return w, h
    except Exception:
        pass
    return None


def get_video_dimensions(path: str):
    """Try ffprobe first, then cv2 as fallback. Returns (w, h) or None."""
    dims = _get_dims_ffprobe(path)
    if dims is None:
        dims = _get_dims_cv2(path)
    return dims


def classify_video(filepath: Path) -> str:
    """Classify one video file. Priority order documented in module docstring."""
    name_lower = filepath.name.lower()

    if name_lower.startswith("screen_recording"):
        return VCAT_SCREEN_REC

    dims = get_video_dimensions(str(filepath))
    if dims is None:
        return VCAT_OTHER

    w, h = dims
    if w == 0 or h == 0:
        return VCAT_OTHER

    ratio      = w / h
    target_169 = 16 / 9
    target_916 = 9  / 16

    if abs(ratio - target_169) / target_169 <= RATIO_TOLERANCE:
        return VCAT_LANDSCAPE
    if abs(ratio - target_916) / target_916 <= RATIO_TOLERANCE:
        return VCAT_PORTRAIT

    return VCAT_OTHER


# ══════════════════════════════════════════════════════════════════
# DOCUMENT constants & logic
# ══════════════════════════════════════════════════════════════════

DCAT_WORD     = "Word Documents"
DCAT_SHEET    = "Spreadsheets"
DCAT_SLIDES   = "Presentations"
DCAT_PDF      = "PDFs"
DCAT_TEXTMD   = "Text & Markdown"
DCAT_CODE     = "Code & Scripts"
DCAT_ARCHIVE  = "Archives"
DCAT_OTHER    = "Other Documents"

DOCUMENT_CATEGORIES = [
    DCAT_WORD, DCAT_SHEET, DCAT_SLIDES, DCAT_PDF,
    DCAT_TEXTMD, DCAT_CODE, DCAT_ARCHIVE, DCAT_OTHER,
]

_DOC_EXT_MAP: dict[str, str] = {}

def _doc_exts(*exts, cat):
    for e in exts:
        _DOC_EXT_MAP[e] = cat

_doc_exts(".doc", ".docx", ".odt", ".rtf", ".dot", ".dotx", ".docm",        cat=DCAT_WORD)
_doc_exts(".xls", ".xlsx", ".ods", ".csv", ".tsv", ".xlsm", ".xlsb",        cat=DCAT_SHEET)
_doc_exts(".ppt", ".pptx", ".odp", ".pot", ".potx", ".pptm",                cat=DCAT_SLIDES)
_doc_exts(".pdf",                                                             cat=DCAT_PDF)
_doc_exts(".txt", ".md", ".rst", ".log", ".nfo", ".readme", ".text",         cat=DCAT_TEXTMD)
_doc_exts(
    ".py", ".js", ".ts", ".jsx", ".tsx", ".html", ".htm", ".css", ".scss",
    ".json", ".xml", ".yaml", ".yml", ".toml", ".ini", ".cfg", ".env",
    ".sql", ".sh", ".bat", ".cmd", ".ps1", ".rb", ".php", ".java", ".c",
    ".cpp", ".cs", ".go", ".rs", ".kt", ".swift", ".r", ".lua", ".pl",
                                                                              cat=DCAT_CODE)
_doc_exts(".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz",
          ".tgz", ".tbz2", ".lz", ".lzma", ".zst", ".cab", ".iso",           cat=DCAT_ARCHIVE)

DOCUMENT_EXTENSIONS = set(_DOC_EXT_MAP.keys()) | {DCAT_OTHER}

def classify_document(filepath: Path) -> str:
    return _DOC_EXT_MAP.get(filepath.suffix.lower(), DCAT_OTHER)

DOCUMENT_CAT_COLORS = {
    DCAT_WORD:    PALETTE["word"],
    DCAT_SHEET:   PALETTE["sheet"],
    DCAT_SLIDES:  PALETTE["slides"],
    DCAT_PDF:     PALETTE["pdf"],
    DCAT_TEXTMD:  PALETTE["textmd"],
    DCAT_CODE:    PALETTE["code"],
    DCAT_ARCHIVE: PALETTE["archive"],
    DCAT_OTHER:   PALETTE["doc_other"],
}

DOCUMENT_CAT_ICONS = {
    DCAT_WORD:    "📝",
    DCAT_SHEET:   "📊",
    DCAT_SLIDES:  "📑",
    DCAT_PDF:     "📄",
    DCAT_TEXTMD:  "📃",
    DCAT_CODE:    "💻",
    DCAT_ARCHIVE: "🗜",
    DCAT_OTHER:   "📁",
}


# ══════════════════════════════════════════════════════════════════
# AUDIO constants & logic
# ══════════════════════════════════════════════════════════════════

ACAT_LOSSLESS = "Lossless"
ACAT_LOSSY    = "Compressed"
ACAT_PLAYLIST = "Playlists"
ACAT_OTHER    = "Other Audio"

AUDIO_CATEGORIES = [ACAT_LOSSLESS, ACAT_LOSSY, ACAT_PLAYLIST, ACAT_OTHER]

_AUDIO_EXT_MAP: dict[str, str] = {}

def _aud_exts(*exts, cat):
    for e in exts:
        _AUDIO_EXT_MAP[e] = cat

_aud_exts(".flac", ".wav", ".aiff", ".aif", ".alac", ".ape",
          ".wv", ".tta", ".dsd", ".dsf", ".dff", ".pcm",       cat=ACAT_LOSSLESS)
_aud_exts(".mp3", ".aac", ".ogg", ".opus", ".m4a", ".wma",
          ".ac3", ".amr", ".mp2", ".mka", ".ra", ".mid",
          ".midi", ".spx", ".voc", ".au",                       cat=ACAT_LOSSY)
_aud_exts(".m3u", ".m3u8", ".pls", ".xspf", ".wpl", ".asx",    cat=ACAT_PLAYLIST)

AUDIO_EXTENSIONS = set(_AUDIO_EXT_MAP.keys())

def classify_audio(filepath: Path) -> str:
    return _AUDIO_EXT_MAP.get(filepath.suffix.lower(), ACAT_OTHER)

AUDIO_CAT_COLORS = {
    ACAT_LOSSLESS: PALETTE["lossless"],
    ACAT_LOSSY:    PALETTE["lossy"],
    ACAT_PLAYLIST: PALETTE["playlist"],
    ACAT_OTHER:    PALETTE["aud_other"],
}

AUDIO_CAT_ICONS = {
    ACAT_LOSSLESS: "🎵",
    ACAT_LOSSY:    "🎧",
    ACAT_PLAYLIST: "📋",
    ACAT_OTHER:    "🔊",
}


# ══════════════════════════════════════════════════════════════════
# Shared helpers
# ══════════════════════════════════════════════════════════════════

def _human_size(nbytes: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if nbytes < 1024:
            return f"{nbytes:.0f} {unit}" if unit == "B" else f"{nbytes:.1f} {unit}"
        nbytes /= 1024
    return f"{nbytes:.1f} TB"


def _safe_folder_name(name: str) -> str:
    """Strip characters Windows forbids in directory names."""
    for ch in r'\/:*?"<>|':
        name = name.replace(ch, "-")
    return name.strip()


def _make_btn(parent, text, cmd, color):
    """Flat styled button with hover effect, shared by both sorter panels."""
    b = tk.Button(
        parent, text=text, command=cmd,
        font=("Segoe UI", 9, "bold"),
        fg=color, bg=PALETTE["card"],
        activeforeground=PALETTE["text"], activebackground=PALETTE["hover"],
        relief="flat", bd=0, cursor="hand2", padx=12, pady=5,
    )
    b.bind("<Enter>", lambda e: b.configure(bg=PALETTE["hover"]))
    b.bind("<Leave>", lambda e: b.configure(bg=PALETTE["card"]))
    return b


def _apply_shared_styles(widget):
    """Apply dark ttk theme. Call once after the first Tk window is created."""
    s = ttk.Style(widget)
    s.theme_use("clam")
    s.configure("Treeview",
                background=PALETTE["panel"], foreground=PALETTE["text"],
                fieldbackground=PALETTE["panel"], borderwidth=0,
                font=("Segoe UI", 9), rowheight=24)
    s.configure("Treeview.Heading",
                background=PALETTE["card"], foreground=PALETTE["subtext"],
                borderwidth=0, font=("Segoe UI", 9, "bold"))
    s.map("Treeview",
          background=[("selected", PALETTE["sel"])],
          foreground=[("selected", PALETTE["text"])])
    s.configure("TNotebook", background=PALETTE["bg"], borderwidth=0)
    s.configure("TNotebook.Tab",
                background=PALETTE["card"], foreground=PALETTE["subtext"],
                padding=[14, 6], font=("Segoe UI", 9, "bold"), borderwidth=0)
    s.map("TNotebook.Tab",
          background=[("selected", PALETTE["panel"])],
          foreground=[("selected", PALETTE["text"])])
    s.configure("TProgressbar",
                troughcolor=PALETTE["card"], background=PALETTE["accent"],
                borderwidth=0, thickness=6)
    s.configure("TScrollbar",
                background=PALETTE["border"], troughcolor=PALETTE["panel"],
                borderwidth=0, arrowsize=0)


# ══════════════════════════════════════════════════════════════════
# Base sorter panel (shared scaffold for both Image and Video tabs)
# ══════════════════════════════════════════════════════════════════

class _SorterPanel(tk.Frame):
    """
    Abstract base for a sorter tab. Subclasses must set:
        self._categories   : list[str]
        self._cat_colors   : dict[str, str]
        self._cat_icons    : dict[str, str]
        self._extensions   : set[str]
    and implement:
        _classify(filepath)  → str   category name
        _scan_row(fp, root)  → tuple  values for the treeview row
        _tree_columns()      → list[tuple(id, heading, width, anchor)]
    """

    def __init__(self, parent):
        super().__init__(parent, bg=PALETTE["bg"])
        # State
        self._folder    = tk.StringVar()
        self._recursive = tk.BooleanVar(value=True)
        self._auto_move = tk.BooleanVar(value=False)
        self._status    = tk.StringVar(value="Choose a folder to begin.")
        self._progress  = tk.DoubleVar(value=0.0)
        self._results: dict[str, list[Path]] = {}
        self._scanning  = False
        self._moving    = False
        self._stop_flag = threading.Event()

    def _init_results(self):
        self._results = {c: [] for c in self._categories}

    # ── Scaffold builder ──────────────────────

    def build(self, subtitle: str):
        """Call from subclass __init__ after setting constants."""
        self._init_results()
        self._build_toolbar(subtitle)
        self._build_options()
        self._build_progress()
        self._build_cards()
        self._build_notebook()
        self._build_action_bar()

    def _build_toolbar(self, subtitle: str):
        picker = tk.Frame(self, bg=PALETTE["card"])
        picker.pack(fill="x", padx=20, pady=(14, 0), ipady=10)

        tk.Label(picker, text="Folder:", font=("Segoe UI", 10, "bold"),
                 fg=PALETTE["subtext"], bg=PALETTE["card"]).pack(side="left", padx=(14, 6))

        tk.Entry(
            picker, textvariable=self._folder,
            font=("Segoe UI", 10), fg=PALETTE["text"],
            bg=PALETTE["panel"], bd=0, insertbackground=PALETTE["text"],
            relief="flat", width=58,
        ).pack(side="left", ipady=5, padx=(0, 8), fill="x", expand=True)

        _make_btn(picker, "Browse",   self._browse,      PALETTE["accent"]).pack(side="left", padx=(0, 6))

        self._scan_btn = _make_btn(picker, "▶  Scan", self._start_scan, PALETTE["accent3"])
        self._scan_btn.pack(side="left", padx=(0, 6))

        self._stop_btn = _make_btn(picker, "■  Stop", self._stop_scan, PALETTE["accent2"])
        self._stop_btn.pack(side="left", padx=(0, 14))
        self._stop_btn.configure(state="disabled")

    def _build_options(self):
        opts = tk.Frame(self, bg=PALETTE["bg"])
        opts.pack(fill="x", padx=20, pady=(8, 0))

        tk.Checkbutton(
            opts, text="Scan sub-folders recursively",
            variable=self._recursive,
            font=("Segoe UI", 9), fg=PALETTE["subtext"],
            bg=PALETTE["bg"], activebackground=PALETTE["bg"],
            activeforeground=PALETTE["text"], selectcolor=PALETTE["card"],
        ).pack(side="left")

        tk.Label(opts, text=" │ ", font=("Segoe UI", 9),
                 fg=PALETTE["border"], bg=PALETTE["bg"]).pack(side="left")

        tk.Checkbutton(
            opts, text="Auto-move files into sub-folders after scan",
            variable=self._auto_move,
            font=("Segoe UI", 9, "bold"), fg=PALETTE["accent2"],
            bg=PALETTE["bg"], activebackground=PALETTE["bg"],
            activeforeground=PALETTE["text"], selectcolor=PALETTE["card"],
        ).pack(side="left")

    def _build_progress(self):
        prog_frame = tk.Frame(self, bg=PALETTE["bg"])
        prog_frame.pack(fill="x", padx=20, pady=(10, 0))

        ttk.Progressbar(
            prog_frame, variable=self._progress,
            maximum=100, mode="determinate",
        ).pack(side="left", fill="x", expand=True)

        tk.Label(
            prog_frame, textvariable=self._status,
            font=("Segoe UI", 9), fg=PALETTE["subtext"],
            bg=PALETTE["bg"], anchor="w",
        ).pack(side="left", padx=(10, 0))

    def _build_cards(self):
        cards = tk.Frame(self, bg=PALETTE["bg"])
        cards.pack(fill="x", padx=20, pady=(14, 0))
        self._count_labels: dict[str, tk.Label] = {}

        for cat in self._categories:
            card = tk.Frame(cards, bg=PALETTE["card"])
            card.pack(side="left", expand=True, fill="x", padx=(0, 8), ipady=10)

            color = self._cat_colors[cat]
            tk.Frame(card, bg=color, width=4).pack(side="left", fill="y")

            inner = tk.Frame(card, bg=PALETTE["card"])
            inner.pack(side="left", padx=12, pady=6)

            title_row = tk.Frame(inner, bg=PALETTE["card"])
            title_row.pack(anchor="w")

            tk.Label(title_row, text=self._cat_icons[cat] + " ",
                     font=("Segoe UI", 11), fg=color, bg=PALETTE["card"]).pack(side="left")
            tk.Label(title_row, text=cat,
                     font=("Segoe UI", 9, "bold"), fg=color, bg=PALETTE["card"]).pack(side="left")

            cnt = tk.Label(inner, text="0 files",
                           font=("Segoe UI", 18, "bold"),
                           fg=PALETTE["text"], bg=PALETTE["card"])
            cnt.pack(anchor="w")
            self._count_labels[cat] = cnt

    def _build_notebook(self):
        nb_frame = tk.Frame(self, bg=PALETTE["bg"])
        nb_frame.pack(fill="both", expand=True, padx=20, pady=(14, 0))

        self._nb = ttk.Notebook(nb_frame)
        self._nb.pack(fill="both", expand=True)

        self._trees: dict[str, ttk.Treeview] = {}
        cols_spec = self._tree_columns()

        for cat in self._categories:
            tab = tk.Frame(self._nb, bg=PALETTE["panel"])
            self._nb.add(tab, text=f"  {self._cat_icons[cat]}  {cat}  ")

            col_ids = [c[0] for c in cols_spec]
            tree = ttk.Treeview(tab, columns=col_ids, show="headings", selectmode="browse")

            for col_id, heading, width, anchor in cols_spec:
                tree.heading(col_id, text=heading)
                tree.column(col_id, width=width, anchor=anchor)

            vsb = ttk.Scrollbar(tab, orient="vertical",   command=tree.yview)
            hsb = ttk.Scrollbar(tab, orient="horizontal", command=tree.xview)
            tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

            vsb.grid(row=0, column=1, sticky="ns")
            hsb.grid(row=1, column=0, sticky="ew")
            tree.grid(row=0, column=0, sticky="nsew")
            tab.rowconfigure(0, weight=1)
            tab.columnconfigure(0, weight=1)

            self._trees[cat] = tree

    def _build_action_bar(self):
        action_bar = tk.Frame(self, bg=PALETTE["panel"])
        action_bar.pack(fill="x", padx=20, pady=(10, 16))

        tk.Label(action_bar, text="Actions:",
                 font=("Segoe UI", 9, "bold"),
                 fg=PALETTE["subtext"], bg=PALETTE["panel"]).pack(side="left", padx=(14, 10))

        self._move_btn = _make_btn(
            action_bar, "Move files into sub-folders",
            lambda: self._organize(move=True), PALETTE["accent2"])
        self._move_btn.pack(side="left", padx=(0, 6))
        self._move_btn.configure(state="disabled")

        self._copy_btn = _make_btn(
            action_bar, "Copy files into sub-folders",
            lambda: self._organize(move=False), PALETTE["accent"])
        self._copy_btn.pack(side="left")
        self._copy_btn.configure(state="disabled")

        tk.Label(action_bar,
                 text="Sub-folders are created inside the scanned folder.",
                 font=("Segoe UI", 8), fg=PALETTE["subtext"], bg=PALETTE["panel"],
                 ).pack(side="right", padx=14)

    # ── Browse / Scan ─────────────────────────

    def _browse(self):
        folder = filedialog.askdirectory(title="Select folder to scan")
        if folder:
            self._folder.set(folder)

    def _start_scan(self):
        folder = self._folder.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("No folder", "Please select a valid folder first.")
            return
        if self._scanning:
            return

        self._init_results()
        for cat in self._categories:
            self._trees[cat].delete(*self._trees[cat].get_children())
            self._count_labels[cat].configure(text="0 files")

        self._stop_flag.clear()
        self._scanning = True
        self._scan_btn.configure(state="disabled")
        self._stop_btn.configure(state="normal")
        self._move_btn.configure(state="disabled")
        self._copy_btn.configure(state="disabled")
        self._progress.set(0)

        threading.Thread(target=self._scan_worker, args=(folder,), daemon=True).start()

    def _stop_scan(self):
        self._stop_flag.set()

    def _scan_worker(self, folder: str):
        try:
            file_type = self._file_type_label()
            self._set_status(f"Collecting {file_type} files…")
            root = Path(folder)

            if self._recursive.get():
                all_files = [p for p in root.rglob("*")
                             if p.is_file() and p.suffix.lower() in self._extensions]
            else:
                all_files = [p for p in root.iterdir()
                             if p.is_file() and p.suffix.lower() in self._extensions]

            total = len(all_files)
            if total == 0:
                self._set_status(f"No {file_type} files found.")
                self.after(0, self._finish_scan)
                return

            self._set_status(f"Found {total} {file_type}(s). Classifying…")

            for i, fp in enumerate(all_files):
                if self._stop_flag.is_set():
                    self.after(0, self._set_status, f"Stopped at {i}/{total}.")
                    break

                self.after(0, self._progress.set, (i / total) * 100)
                self.after(0, self._set_status, f"[{i+1}/{total}]  {fp.name}")

                try:
                    cat     = self._classify(fp)
                    row_vals = self._scan_row(fp, root)
                except Exception:
                    cat      = self._categories[-1]   # fallback: last cat = "Other"
                    row_vals = self._error_row(fp, root)

                self._results[cat].append(fp)
                self.after(0, self._add_row, cat, row_vals)
                self.after(0, self._refresh_count, cat)

            self.after(0, self._progress.set, 100)
            found = sum(len(v) for v in self._results.values())
            self.after(0, self._set_status, f"Done — {found} {file_type}(s) classified.")

        except Exception as exc:
            self.after(0, self._set_status, f"Error: {exc}")
        finally:
            self.after(0, self._finish_scan)

    def _add_row(self, cat: str, vals: tuple):
        tree = self._trees[cat]
        iid = tree.insert("", "end", values=vals)
        tag = "even" if len(tree.get_children()) % 2 == 0 else "odd"
        tree.item(iid, tags=(tag,))
        tree.tag_configure("odd",  background=PALETTE["panel"])
        tree.tag_configure("even", background=PALETTE["card"])

    def _refresh_count(self, cat: str):
        n = len(self._results[cat])
        self._count_labels[cat].configure(text=f"{n} file{'s' if n != 1 else ''}")

    def _finish_scan(self):
        self._scanning = False
        self._scan_btn.configure(state="normal")
        self._stop_btn.configure(state="disabled")
        has = any(self._results[c] for c in self._categories)
        st  = "normal" if has else "disabled"
        self._move_btn.configure(state=st)
        self._copy_btn.configure(state=st)
        if has and self._auto_move.get() and not self._stop_flag.is_set():
            self.after(200, lambda: self._organize(move=True, silent=True))

    # ── Organize ──────────────────────────────

    def _organize(self, move: bool, silent: bool = False):
        folder = self._folder.get().strip()
        if not folder or self._moving:
            return

        total = sum(len(v) for v in self._results.values())
        if total == 0:
            if not silent:
                messagebox.showinfo("Nothing to do", "No files to organise.")
            return

        verb = "move" if move else "copy"
        if not silent:
            preview = "\n".join(
                f"  {self._cat_icons[c]}  {c}/  ({len(self._results[c])} files)"
                for c in self._categories if self._results[c]
            )
            if not messagebox.askyesno(
                "Confirm",
                f"This will {verb} {total} file(s) into sub-folders inside:\n{folder}\n\n"
                f"Destination folders:\n{preview}\n\nProceed?"
            ):
                return

        self._moving = True
        self._move_btn.configure(state="disabled")
        self._copy_btn.configure(state="disabled")
        self._scan_btn.configure(state="disabled")
        self._progress.set(0)

        threading.Thread(target=self._move_worker, args=(folder, move), daemon=True).start()

    def _move_worker(self, folder: str, move: bool):
        root   = Path(folder)
        verb   = "Moving" if move else "Copying"
        errors: list[str] = []
        done   = 0

        work: list[tuple[Path, Path]] = []
        for cat, files in self._results.items():
            if not files:
                continue
            dest_dir = root / _safe_folder_name(cat)
            try:
                dest_dir.mkdir(exist_ok=True)
            except Exception as e:
                errors.append(f"mkdir {cat}: {e}")
                continue

            for fp in files:
                dest = dest_dir / fp.name
                if dest.exists() and dest.resolve() != fp.resolve():
                    stem, suffix = fp.stem, fp.suffix
                    n = 1
                    while dest.exists():
                        dest = dest_dir / f"{stem}_{n}{suffix}"
                        n += 1
                work.append((fp, dest))

        actual = len(work)
        for i, (src, dst) in enumerate(work):
            self.after(0, self._progress.set, (i / actual * 100) if actual else 100)
            self.after(0, self._set_status,
                       f"{verb} [{i+1}/{actual}]  {src.name}  →  {dst.parent.name}/")
            try:
                if src.resolve() == dst.resolve():
                    done += 1
                    continue
                if move:
                    shutil.move(str(src), str(dst))
                else:
                    shutil.copy2(str(src), str(dst))
                done += 1
            except Exception as e:
                errors.append(f"{src.name}: {e}")

        self.after(0, self._progress.set, 100)
        action = "moved" if move else "copied"
        lines  = [f"✓  {done} file(s) {action} successfully.\n"]
        for cat in self._categories:
            n = len(self._results[cat])
            if n:
                lines.append(f"  {self._cat_icons[cat]}  {_safe_folder_name(cat)}/  ← {n} file(s)")
        if errors:
            lines.append(f"\n✗  {len(errors)} error(s):")
            lines.extend(f"   • {e}" for e in errors[:15])
            if len(errors) > 15:
                lines.append(f"   … and {len(errors)-15} more.")

        self.after(0, self._set_status, f"Done — {done}/{actual} files {action}.")
        self.after(0, messagebox.showinfo, "Done", "\n".join(lines))
        self.after(0, self._finish_move)

    def _finish_move(self):
        self._moving = False
        self._scan_btn.configure(state="normal")
        has = any(self._results[c] for c in self._categories)
        st  = "normal" if has else "disabled"
        self._move_btn.configure(state=st)
        self._copy_btn.configure(state=st)

    # ── Helpers ───────────────────────────────

    def _set_status(self, msg: str):
        self._status.set(msg)

    def _file_type_label(self) -> str:
        return "file"

    def _error_row(self, fp: Path, root: Path) -> tuple:
        rel = str(fp.parent.relative_to(root)) if fp.parent != root else "."
        return (fp.name, fp.suffix.lower(), rel, "—")

    # ── Abstract interface ────────────────────

    def _classify(self, filepath: Path) -> str:
        raise NotImplementedError

    def _scan_row(self, fp: Path, root: Path) -> tuple:
        raise NotImplementedError

    def _tree_columns(self) -> list:
        raise NotImplementedError


# ══════════════════════════════════════════════════════════════════
# Image Sorter panel
# ══════════════════════════════════════════════════════════════════

class ImageSorterPanel(_SorterPanel):

    def __init__(self, parent):
        super().__init__(parent)
        self._categories  = IMAGE_CATEGORIES
        self._cat_colors  = IMAGE_CAT_COLORS
        self._cat_icons   = IMAGE_CAT_ICONS
        self._extensions  = IMAGE_EXTENSIONS
        self.build(subtitle="wallpapers · screenshots · webp · square · everything else")

    def _file_type_label(self):
        return "image"

    def _classify(self, filepath: Path) -> str:
        return classify_image(filepath)

    def _tree_columns(self):
        return [
            ("name",   "File Name", 300, "w"),
            ("ext",    "Ext",        60, "center"),
            ("folder", "Folder",    340, "w"),
            ("size",   "Size",       80, "e"),
        ]

    def _scan_row(self, fp: Path, root: Path) -> tuple:
        size_str = _human_size(fp.stat().st_size)
        rel      = str(fp.parent.relative_to(root)) if fp.parent != root else "."
        return (fp.name, fp.suffix.lower(), rel, size_str)

    def _error_row(self, fp: Path, root: Path) -> tuple:
        rel = str(fp.parent.relative_to(root)) if fp.parent != root else "."
        return (fp.name, fp.suffix.lower(), rel, "—")


# ══════════════════════════════════════════════════════════════════
# Video Sorter panel
# ══════════════════════════════════════════════════════════════════

class VideoSorterPanel(_SorterPanel):

    def __init__(self, parent):
        super().__init__(parent)
        self._categories  = VIDEO_CATEGORIES
        self._cat_colors  = VIDEO_CAT_COLORS
        self._cat_icons   = VIDEO_CAT_ICONS
        self._extensions  = VIDEO_EXTENSIONS
        self.build(subtitle="screen recordings · 16:9 · 9:16 · everything else")

    def _file_type_label(self):
        return "video"

    def _classify(self, filepath: Path) -> str:
        return classify_video(filepath)

    def _tree_columns(self):
        return [
            ("name",       "File Name",  280, "w"),
            ("folder",     "Folder",     320, "w"),
            ("resolution", "Resolution", 110, "center"),
            ("ratio",      "Ratio",       90, "center"),
        ]

    def _scan_row(self, fp: Path, root: Path) -> tuple:
        dims = get_video_dimensions(str(fp))
        if dims:
            w, h = dims
            res_str   = f"{w}×{h}"
            gcd       = math.gcd(w, h)
            ratio_str = f"{w//gcd}:{h//gcd}"
        else:
            res_str   = "unknown"
            ratio_str = "—"
        rel = str(fp.parent.relative_to(root)) if fp.parent != root else "."
        return (fp.name, rel, res_str, ratio_str)

    def _error_row(self, fp: Path, root: Path) -> tuple:
        rel = str(fp.parent.relative_to(root)) if fp.parent != root else "."
        return (fp.name, rel, "error", "—")


# ══════════════════════════════════════════════════════════════════
# Document Sorter panel
# ══════════════════════════════════════════════════════════════════

class DocumentSorterPanel(_SorterPanel):

    def __init__(self, parent):
        super().__init__(parent)
        self._categories = DOCUMENT_CATEGORIES
        self._cat_colors = DOCUMENT_CAT_COLORS
        self._cat_icons  = DOCUMENT_CAT_ICONS
        self._extensions = DOCUMENT_EXTENSIONS
        self.build(subtitle="word · spreadsheets · presentations · pdf · text · code · archives")

    def _file_type_label(self):
        return "document"

    def _classify(self, filepath: Path) -> str:
        return classify_document(filepath)

    def _tree_columns(self):
        return [
            ("name",   "File Name", 320, "w"),
            ("ext",    "Ext",        70, "center"),
            ("folder", "Folder",    300, "w"),
            ("size",   "Size",       80, "e"),
        ]

    def _scan_row(self, fp: Path, root: Path) -> tuple:
        size_str = _human_size(fp.stat().st_size)
        rel      = str(fp.parent.relative_to(root)) if fp.parent != root else "."
        return (fp.name, fp.suffix.lower(), rel, size_str)

    def _error_row(self, fp: Path, root: Path) -> tuple:
        rel = str(fp.parent.relative_to(root)) if fp.parent != root else "."
        return (fp.name, fp.suffix.lower(), rel, "—")


# ══════════════════════════════════════════════════════════════════
# Audio Sorter panel
# ══════════════════════════════════════════════════════════════════

class AudioSorterPanel(_SorterPanel):

    def __init__(self, parent):
        super().__init__(parent)
        self._categories = AUDIO_CATEGORIES
        self._cat_colors = AUDIO_CAT_COLORS
        self._cat_icons  = AUDIO_CAT_ICONS
        self._extensions = AUDIO_EXTENSIONS
        self.build(subtitle="lossless · compressed · playlists · everything else")

    def _file_type_label(self):
        return "audio"

    def _classify(self, filepath: Path) -> str:
        return classify_audio(filepath)

    def _tree_columns(self):
        return [
            ("name",   "File Name", 340, "w"),
            ("ext",    "Ext",        70, "center"),
            ("folder", "Folder",    320, "w"),
            ("size",   "Size",       80, "e"),
        ]

    def _scan_row(self, fp: Path, root: Path) -> tuple:
        size_str = _human_size(fp.stat().st_size)
        rel      = str(fp.parent.relative_to(root)) if fp.parent != root else "."
        return (fp.name, fp.suffix.lower(), rel, size_str)

    def _error_row(self, fp: Path, root: Path) -> tuple:
        rel = str(fp.parent.relative_to(root)) if fp.parent != root else "."
        return (fp.name, fp.suffix.lower(), rel, "—")


# ══════════════════════════════════════════════════════════════════
# Main application window
# ══════════════════════════════════════════════════════════════════

class MediaSortApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("MediaSort")
        self.geometry("1080x780")
        self.minsize(840, 600)
        self.configure(bg=PALETTE["bg"])

        _apply_shared_styles(self)
        self._build_ui()

    def _build_ui(self):
        # ── App header ───────────────────────────
        hdr = tk.Frame(self, bg=PALETTE["bg"])
        hdr.pack(fill="x", padx=20, pady=(18, 0))

        tk.Label(
            hdr, text="MediaSort",
            font=("Segoe UI", 22, "bold"),
            fg=PALETTE["accent"], bg=PALETTE["bg"],
        ).pack(side="left")

        tk.Label(
            hdr, text=" — images, videos, documents & audio, organised",
            font=("Segoe UI", 11),
            fg=PALETTE["subtext"], bg=PALETTE["bg"],
        ).pack(side="left", pady=(8, 0))

        # ── Mode switcher ────────────────────────
        switcher = tk.Frame(self, bg=PALETTE["bg"])
        switcher.pack(fill="x", padx=20, pady=(12, 0))

        self._mode = tk.StringVar(value="images")

        _tab_defs = [
            ("Images",    "images",    "🖼"),
            ("Videos",    "videos",    "🎬"),
            ("Documents", "documents", "📄"),
            ("Audio",     "audio",     "🎵"),
        ]
        self._tab_btns: dict[str, tk.Button] = {}

        def _select_tab(value):
            self._mode.set(value)
            _refresh_tabs()

        for text, value, icon in _tab_defs:
            b = tk.Button(
                switcher, text=f"  {icon}  {text}  ",
                command=lambda v=value: _select_tab(v),
                font=("Segoe UI", 10, "bold"),
                fg=PALETTE["text"], bg=PALETTE["card"],
                activeforeground=PALETTE["text"], activebackground=PALETTE["hover"],
                relief="flat", bd=0, cursor="hand2", padx=8, pady=6,
            )
            b.pack(side="left", padx=(0, 4))
            self._tab_btns[value] = b

        sep = tk.Frame(self, bg=PALETTE["border"], height=1)
        sep.pack(fill="x", padx=20, pady=(10, 0))

        # ── Panels ───────────────────────────────
        self._img_panel  = ImageSorterPanel(self)
        self._vid_panel  = VideoSorterPanel(self)
        self._doc_panel  = DocumentSorterPanel(self)
        self._aud_panel  = AudioSorterPanel(self)

        self._panels = {
            "images":    self._img_panel,
            "videos":    self._vid_panel,
            "documents": self._doc_panel,
            "audio":     self._aud_panel,
        }

        def _refresh_tabs():
            mode = self._mode.get()
            for key, panel in self._panels.items():
                panel.pack_forget()
            self._panels[mode].pack(fill="both", expand=True)
            for key, btn in self._tab_btns.items():
                if key == mode:
                    btn.configure(bg=PALETTE["accent"], fg=PALETTE["bg"])
                else:
                    btn.configure(bg=PALETTE["card"], fg=PALETTE["text"])

        _refresh_tabs()   # show images panel by default


# ══════════════════════════════════════════════════════════════════
# Entry point
# ══════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = MediaSortApp()
    app.mainloop()
