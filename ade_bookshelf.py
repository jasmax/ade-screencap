#!/usr/bin/env python3

from __future__ import annotations

import os
import sys
import io
import time
import hashlib
import ctypes
import ctypes.wintypes
import platform
import subprocess
import argparse
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Optional

# ── Platform guard ────────────────────────────────────────────────────────────
if platform.system() != "Windows":
    print("ERROR: This script uses Windows-native capture APIs and must run on Windows.",
          file=sys.stderr)
    sys.exit(1)

# ── Virtual-environment guard ─────────────────────────────────────────────────
def _in_venv() -> bool:
    return (
        hasattr(sys, "real_prefix")
        or (hasattr(sys, "base_prefix") and sys.base_prefix != sys.prefix)
    )

if not _in_venv():
    _Y = "\033[93m" if sys.stderr.isatty() else ""
    _R = "\033[0m"  if sys.stderr.isatty() else ""
    print(f"\n{_Y}  ⚠  WARNING: Not running inside a virtual environment.{_R}\n"
          f"     Use the launcher:  python run_ade.py\n", file=sys.stderr)

# ── Optional colour output ─────────────────────────────────────────────────────
try:
    from colorama import init as colorama_init, Fore, Style
    colorama_init(autoreset=True)
    HAS_COLOR = True
except ImportError:
    HAS_COLOR = False
    class _Dummy:
        def __getattr__(self, _): return ""
    Fore = Style = _Dummy()

# ── Capture dependencies ───────────────────────────────────────────────────────
try:
    import mss
    import mss.tools
    HAS_MSS = True
except ImportError:
    HAS_MSS = False

try:
    from PIL import Image
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as rl_canvas
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False


# ═══════════════════════════════════════════════════════════════════════════════
# 1.  CONSOLE HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def _color(text: str, color: str) -> str:
    return f"{color}{text}{Style.RESET_ALL}" if HAS_COLOR else text

def header(msg: str) -> None:
    print(_color(f"\n{'═'*60}", Fore.CYAN))
    print(_color(f"  {msg}", Fore.CYAN + Style.BRIGHT))
    print(_color(f"{'═'*60}", Fore.CYAN))

def info(msg: str)    -> None: print(_color(f"  ℹ  {msg}", Fore.BLUE))
def success(msg: str) -> None: print(_color(f"  ✔  {msg}", Fore.GREEN))
def warn(msg: str)    -> None: print(_color(f"  ⚠  {msg}", Fore.YELLOW))
def error(msg: str)   -> None: print(_color(f"  ✖  {msg}", Fore.RED))

def prompt(msg: str) -> str:
    return input(_color(f"\n  → {msg}: ", Fore.MAGENTA)).strip()


# ═══════════════════════════════════════════════════════════════════════════════
# 2.  PATH NORMALISATION
# ═══════════════════════════════════════════════════════════════════════════════

def normalize_path(raw: str) -> Path:
    """Strip stray quotes/spaces and resolve to an absolute Path."""
    return Path(raw.strip().strip("\"'")).expanduser().resolve()


# ═══════════════════════════════════════════════════════════════════════════════
# 3.  WIN32 WINDOW MANAGEMENT  (ctypes — no extra packages)
# ═══════════════════════════════════════════════════════════════════════════════

user32   = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

SW_RESTORE  = 9
SW_MAXIMIZE = 3
WM_KEYDOWN  = 0x0100
WM_KEYUP    = 0x0101
VK_RIGHT    = 0x27          # Right-Arrow virtual key code

# Keywords that identify the ADE window title
ADE_TITLE_KEYWORDS = ["adobe digital editions", "digital editions"]


def _enum_windows() -> list[int]:
    """Return HWNDs of all visible top-level windows."""
    hwnds: list[int] = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
    def _cb(hwnd, _):
        if user32.IsWindowVisible(hwnd):
            hwnds.append(hwnd)
        return True

    user32.EnumWindows(_cb, 0)
    return hwnds


def _get_window_title(hwnd: int) -> str:
    buf = ctypes.create_unicode_buffer(512)
    user32.GetWindowTextW(hwnd, buf, 512)
    return buf.value


def find_ade_hwnd() -> Optional[int]:
    """Return the HWND of the first window whose title matches ADE keywords."""
    for hwnd in _enum_windows():
        title = _get_window_title(hwnd).lower()
        if any(kw in title for kw in ADE_TITLE_KEYWORDS):
            return hwnd
    return None


def wait_for_ade_hwnd(timeout: float = 30.0, poll: float = 0.4) -> Optional[int]:
    """Poll until the ADE window appears. Returns HWND or None on timeout."""
    info(f"Waiting up to {int(timeout)}s for ADE window…")
    deadline = time.time() + timeout
    while time.time() < deadline:
        hwnd = find_ade_hwnd()
        if hwnd:
            return hwnd
        time.sleep(poll)
    return None


def get_window_rect(hwnd: int) -> Optional[tuple[int, int, int, int]]:
    """
    Return the full window rect (left, top, width, height) including chrome.
    Used only for focus/maximise checks — NOT for capture.
    """
    rect = ctypes.wintypes.RECT()
    if user32.GetWindowRect(hwnd, ctypes.byref(rect)):
        w = rect.right  - rect.left
        h = rect.bottom - rect.top
        if w > 10 and h > 10:
            return (rect.left, rect.top, w, h)
    return None


def get_client_rect_screen(hwnd: int) -> Optional[tuple[int, int, int, int]]:
    """
    Return the CLIENT area rect in screen coordinates (left, top, width, height).

    GetClientRect gives the client area in window-local coords (always 0,0 origin).
    ClientToScreen translates the top-left corner to screen coords.
    This excludes the title bar, menu bar, and window border — but still includes
    ADE's own toolbar and navigation panels inside the client area.
    """
    client_rect = ctypes.wintypes.RECT()
    if not user32.GetClientRect(hwnd, ctypes.byref(client_rect)):
        return None

    w = client_rect.right  - client_rect.left
    h = client_rect.bottom - client_rect.top
    if w < 10 or h < 10:
        return None

    # Translate client (0,0) → screen coordinates
    pt = ctypes.wintypes.POINT(0, 0)
    user32.ClientToScreen(hwnd, ctypes.byref(pt))

    return (pt.x, pt.y, w, h)


def focus_and_maximise(hwnd: int) -> None:
    """
    Restore → foreground → maximise the ADE window using Win32 calls.
    Using ShowWindow + SetForegroundWindow is more reliable than
    pygetwindow on multi-monitor / high-DPI setups.
    """
    try:
        user32.ShowWindow(hwnd, SW_RESTORE)
        time.sleep(0.3)
        user32.SetForegroundWindow(hwnd)
        time.sleep(0.3)
        user32.ShowWindow(hwnd, SW_MAXIMIZE)
        time.sleep(0.8)
        user32.SetForegroundWindow(hwnd)   # re-foreground after maximise
    except Exception as exc:
        warn(f"focus_and_maximise: {exc}")

    # Let ADE fully paint the maximised frame before first capture
    time.sleep(2.0)


def send_next_page(hwnd: int) -> None:
    """
    Send a Right-Arrow keypress directly to the ADE window via
    PostMessage — no pyautogui, no focus required.
    PostMessage is asynchronous and does not steal focus.
    """
    user32.PostMessageW(hwnd, WM_KEYDOWN, VK_RIGHT, 0)
    time.sleep(0.05)
    user32.PostMessageW(hwnd, WM_KEYUP,   VK_RIGHT, 0)


# ═══════════════════════════════════════════════════════════════════════════════
# 4.  PAGE CROP  — fixed offset crop
# ═══════════════════════════════════════════════════════════════════════════════

# Measured from the 3840×2160 ADE window screenshot:
#
#   Full capture : 3840 × 2160 px
#   Top offset   :  125 px  (ADE toolbar + menu bar)
#   Left offset  :  516 px  (grey left margin)
#   Crop size    : 2794 × 1962 px
#   Bottom check : 125 + 1962 = 2087 → bottom margin = 2160 - 2087 = 73 px
#   Right check  : 516 + 2794 = 3310 → right  margin = 3840 - 3310 = 530 px
#
# These are the defaults; --crop-w / --crop-h / --top / --left override them.

DEFAULT_CROP_W  = 2794
DEFAULT_CROP_H  = 1962
DEFAULT_TOP     = 125     # px from top of client area to start of page content
DEFAULT_LEFT    = 516     # px from left of client area


def offset_crop(img:    "Image.Image",
                top:    int = DEFAULT_TOP,
                left:   int = DEFAULT_LEFT,
                crop_w: int = DEFAULT_CROP_W,
                crop_h: int = DEFAULT_CROP_H) -> "Image.Image":
    """
    Crop *crop_w* × *crop_h* pixels starting at (*left*, *top*) from *img*.

    This removes ADE's toolbar (top *top* px), grey left/right margins
    (*left* px each side), and the page-counter bar at the bottom.

    Layout at 3840×2160
    ───────────────────
      top    = 125 px  → trims ADE toolbar + menu row
      left   = 516 px  → trims grey left margin
      width  = 2794 px → page content width
      height = 1962 px → page content height
      bottom = 125 + 1962 = 2087 → bottom margin = 2160 - 2087 = 73 px
      right  = 516 + 2794 = 3310 → right  margin = 3840 - 3310 = 530 px

    If the crop box would exceed the image dimensions it is clamped
    automatically so the function never raises an error on smaller displays.
    """
    W, H   = img.size
    l      = max(0, min(left,  W))
    t      = max(0, min(top,   H))
    r      = max(0, min(l + crop_w, W))
    b      = max(0, min(t + crop_h, H))
    return img.crop((l, t, r, b))


# ═══════════════════════════════════════════════════════════════════════════════
# 5.  MSS CAPTURE  (Windows DXGI / BitBlt)
# ═══════════════════════════════════════════════════════════════════════════════

# Black-frame detection
BLACK_THRESHOLD    = 15    # mean brightness 0–255; below = black
RENDER_TIMEOUT     = 10.0  # seconds to wait for a non-black frame
RENDER_POLL        = 0.25  # seconds between render polls
END_OF_BOOK_STREAK = 5     # consecutive identical hashes → end of book

# First-page stability: wait until hash stops changing for this many polls
STABLE_HASH_COUNT  = 3     # consecutive identical hashes = page fully rendered
STABLE_POLL        = 0.4   # seconds between stability polls
STABLE_TIMEOUT     = 20.0  # max seconds to wait for first page to stabilise


def _capture_client(hwnd: int) -> Optional["Image.Image"]:
    """
    Capture only the CLIENT area of the ADE window using mss (DXGI BitBlt).
    Client area = window interior excluding title bar and window border.
    Returns a PIL RGB Image, or None on failure.
    """
    rect = get_client_rect_screen(hwnd)
    if rect is None:
        return None
    left, top, width, height = rect

    try:
        with mss.mss() as sct:
            monitor = {"left": left, "top": top, "width": width, "height": height}
            raw = sct.grab(monitor)
            img = Image.frombytes("RGB", raw.size, raw.bgra, "raw", "BGRX")
            return img
    except Exception as exc:
        warn(f"mss capture failed: {exc}")
        return None


def _brightness(img: "Image.Image") -> float:
    """Mean brightness of a 32×32 greyscale thumbnail (0–255)."""
    small = img.resize((32, 32)).convert("L")
    return sum(small.getdata()) / (32 * 32)


def _img_hash(img: "Image.Image") -> str:
    """64×64 greyscale MD5 — fast perceptual hash for duplicate detection."""
    small = img.resize((64, 64)).convert("L")
    return hashlib.md5(small.tobytes()).hexdigest()


def _is_black(img: "Image.Image") -> bool:
    return _brightness(img) < BLACK_THRESHOLD


def _wait_for_render(hwnd: int,
                     timeout: float = RENDER_TIMEOUT,
                     poll:    float = RENDER_POLL) -> Optional["Image.Image"]:
    """
    Poll mss until the ADE client area shows a non-black frame.
    Returns the first bright frame, or the last captured frame on timeout.
    """
    deadline = time.time() + timeout
    last_img = None

    while time.time() < deadline:
        img = _capture_client(hwnd)
        if img is not None:
            last_img = img
            if _brightness(img) > BLACK_THRESHOLD:
                return img
        time.sleep(poll)

    return last_img


def _wait_for_stable_page(hwnd: int,
                           timeout: float = STABLE_TIMEOUT,
                           poll:    float = STABLE_POLL,
                           streak:  int   = STABLE_HASH_COUNT
                           ) -> Optional["Image.Image"]:
    """
    Wait until the ADE page content has FULLY loaded and stopped changing.

    ADE renders pages progressively — text and images appear gradually.
    Grabbing too early gives a partially-rendered page (the 'quarter page' bug).

    This function polls until the same content hash is observed *streak* times
    in a row, meaning the page has stopped updating.

    Returns the stable frame, or the last frame captured on timeout.
    """
    info(f"Waiting for page to fully render (up to {int(timeout)}s)…")
    deadline     = time.time() + timeout
    prev_hash    = ""
    match_count  = 0
    last_img     = None

    while time.time() < deadline:
        img = _capture_client(hwnd)
        if img is None:
            time.sleep(poll)
            continue

        if _is_black(img):
            time.sleep(poll)
            continue

        last_img = img
        h = _img_hash(img)

        if h == prev_hash:
            match_count += 1
            if match_count >= streak:
                success(f"Page stable after {match_count} identical frames.")
                return img
        else:
            match_count = 0
            prev_hash   = h

        time.sleep(poll)

    warn("Stability timeout — using last captured frame.")
    return last_img


# ═══════════════════════════════════════════════════════════════════════════════
# 5.  JPEG COMPRESSION
# ═══════════════════════════════════════════════════════════════════════════════

def _compress_jpeg(img: "Image.Image", quality: int = 85) -> bytes:
    """
    Compress a PIL Image to JPEG bytes in memory.
    JPEG gives 5–15× smaller files than PNG for book page screenshots,
    with negligible visual quality loss at quality=85.
    """
    buf = io.BytesIO()
    # Convert RGBA → RGB (JPEG does not support alpha)
    rgb = img.convert("RGB")
    rgb.save(buf, format="JPEG", quality=quality, optimize=True)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
# 6.  PDF ASSEMBLY
# ═══════════════════════════════════════════════════════════════════════════════

def _build_pdf(jpeg_frames: list[bytes], output_path: Path, dpi: int = 150) -> None:
    """
    Assemble JPEG-compressed frames into a single PDF using Pillow's
    built-in multi-page PDF writer.

    Each page is sized to match the image's pixel dimensions at *dpi*,
    so the PDF renders at the original capture resolution.

    Falls back to reportlab if the Pillow PDF writer is unavailable.
    """
    if not jpeg_frames:
        raise ValueError("No frames to assemble into PDF.")

    # ── Primary: Pillow multi-page PDF ────────────────────────────────────────
    try:
        images = [Image.open(io.BytesIO(j)).convert("RGB") for j in jpeg_frames]
        first  = images[0]
        rest   = images[1:]

        # Pillow saves multi-page PDF with append_images
        first.save(
            str(output_path),
            format="PDF",
            save_all=True,
            append_images=rest,
            resolution=dpi,
        )
        return
    except Exception as e:
        warn(f"Pillow PDF writer failed ({e}) — falling back to reportlab.")

    # ── Fallback: reportlab ────────────────────────────────────────────────────
    if not HAS_REPORTLAB:
        raise RuntimeError("reportlab not available and Pillow PDF failed.")

    import tempfile
    a4_w, a4_h = A4
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        c   = rl_canvas.Canvas(str(output_path), pagesize=A4)
        for i, jpeg_bytes in enumerate(jpeg_frames):
            img_path = tmp / f"p{i:05d}.jpg"
            img_path.write_bytes(jpeg_bytes)
            img = Image.open(str(img_path))
            iw, ih  = img.size
            ratio   = min(a4_w / iw, a4_h / ih)
            draw_w  = iw * ratio
            draw_h  = ih * ratio
            x       = (a4_w - draw_w) / 2
            y       = (a4_h - draw_h) / 2
            c.drawImage(str(img_path), x, y, width=draw_w, height=draw_h)
            c.showPage()
        c.save()


# ═══════════════════════════════════════════════════════════════════════════════
# 7.  AUTO-DETECT ADE LIBRARY
# ═══════════════════════════════════════════════════════════════════════════════

def detect_ade_library() -> Optional[Path]:
    user = Path.home()
    candidates = [
        user / "Documents"          / "My Digital Editions",
        user / "OneDrive"           / "Documents" / "My Digital Editions",
        user / "OneDrive"           / "Documentos" / "My Digital Editions",
        user / "My Documents"       / "My Digital Editions",
    ]
    return next((c for c in candidates if c.exists()), None)


# ═══════════════════════════════════════════════════════════════════════════════
# 8.  EPUB / PDF METADATA
# ═══════════════════════════════════════════════════════════════════════════════

_DC_NS        = "http://purl.org/dc/elements/1.1/"
_CONTAINER_NS = {"c": "urn:oasis:names:tc:opendocument:xmlns:container"}

def _opf_text(root, tag: str) -> str:
    el = root.find(f".//{{{_DC_NS}}}{tag}")
    return el.text.strip() if el is not None and el.text else "Unknown"



def _base_meta(path: Path, fmt: str) -> dict:
    return {
        "title":    path.stem.replace("_", " ").replace("-", " ").title(),
        "author":   "Unknown", "publisher": "Unknown",
        "date":     "Unknown", "language":  "Unknown",
        "format":   fmt,       "path":      path,
        "size_mb":  round(path.stat().st_size / 1_048_576, 2),
        "modified": datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d"),
    }

def epub_metadata(path: Path) -> dict:
    meta = _base_meta(path, "EPUB")
    try:
        with zipfile.ZipFile(path, "r") as z:
            container = ET.fromstring(z.read("META-INF/container.xml"))
            opf_path  = container.find(".//c:rootfile", _CONTAINER_NS).get("full-path")
            root      = ET.fromstring(z.read(opf_path))
            meta["title"]     = _opf_text(root, "title")
            meta["author"]    = _opf_text(root, "creator")
            meta["publisher"] = _opf_text(root, "publisher")
            meta["date"]      = _opf_text(root, "date")
            meta["language"]  = _opf_text(root, "language")
    except Exception:
        pass
    return meta

def pdf_metadata(path: Path) -> dict:
    return _base_meta(path, "PDF")

def book_metadata(p: Path) -> dict:
    if p.suffix.lower() == ".epub": return epub_metadata(p)
    if p.suffix.lower() == ".pdf":  return pdf_metadata(p)
    return {}


def _base_meta(path: Path, fmt: str) -> dict:
    return {
        "title":      path.stem.replace("_", " ").replace("-", " ").title(),
        "author":     "Unknown", "publisher": "Unknown",
        "date":       "Unknown", "language":  "Unknown",
        "format":     fmt,       "path":      path,
        "size_mb":    round(path.stat().st_size / 1_048_576, 2),
        "modified":   datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d"),
    }

def epub_metadata(path: Path) -> dict:
    meta = _base_meta(path, "EPUB")
    try:
        with zipfile.ZipFile(path, "r") as z:
            container = ET.fromstring(z.read("META-INF/container.xml"))
            opf_path  = container.find(".//c:rootfile", _CONTAINER_NS).get("full-path")
            root      = ET.fromstring(z.read(opf_path))
            meta["title"]     = _opf_text(root, "title")
            meta["author"]    = _opf_text(root, "creator")
            meta["publisher"] = _opf_text(root, "publisher")
            meta["date"]      = _opf_text(root, "date")
            meta["language"]  = _opf_text(root, "language")
    except Exception:
        pass
    return meta

def pdf_metadata(path: Path) -> dict:
    return _base_meta(path, "PDF")

def book_metadata(p: Path) -> dict:
    if p.suffix.lower() == ".epub": return epub_metadata(p)
    if p.suffix.lower() == ".pdf":  return pdf_metadata(p)
    return {}


# ═══════════════════════════════════════════════════════════════════════════════
# 9.  LIBRARY SCANNER
# ═══════════════════════════════════════════════════════════════════════════════

class ADELibrary:
    SUPPORTED = {".epub", ".pdf"}

    def __init__(self, library_path: Path):
        self.root    = library_path
        self.books:  list[dict]             = []
        self.shelves: dict[str, list[dict]] = {}
        self._scan()

    def _scan(self) -> None:
        files = sorted(
            p for p in self.root.rglob("*")
            if p.suffix.lower() in self.SUPPORTED and p.is_file()
        )
        self.books = [book_metadata(f) for f in files]
        self._build_shelves()

    def _build_shelves(self) -> None:
        manifest = self.root / "manifest.xml"
        self.shelves = {}
        if manifest.exists():
            self._parse_manifest(manifest)
        else:
            for book in self.books:
                shelf = book["path"].parent.name
                if shelf == self.root.name:
                    shelf = "All Books"
                self.shelves.setdefault(shelf, []).append(book)
        assigned = {b["path"] for shelf in self.shelves.values() for b in shelf}
        unassigned = [b for b in self.books if b["path"] not in assigned]
        if unassigned:
            self.shelves.setdefault("Unshelved", []).extend(unassigned)

    def _parse_manifest(self, mp: Path) -> None:
        try:
            root = ET.parse(mp).getroot()
            idx  = {b["path"]: b for b in self.books}
            for shelf_el in root.findall(".//bookshelf"):
                name, books = shelf_el.get("name", "Unnamed"), []
                for item in shelf_el.findall("item"):
                    cand = (self.root / item.get("src", "")).resolve()
                    if cand in idx:
                        books.append(idx[cand])
                    else:
                        fname = Path(item.get("src", "")).name
                        books.extend(b for b in self.books if b["path"].name == fname)
                if books:
                    self.shelves[name] = books
        except ET.ParseError:
            self._build_shelves()

    def search(self, q: str) -> list[dict]:
        q = q.lower()
        return [b for b in self.books
                if q in b["title"].lower() or q in b["author"].lower()]

    def shelf_names(self) -> list[str]:
        return sorted(self.shelves.keys())


# ═══════════════════════════════════════════════════════════════════════════════
# 10. BOOK OPENER
# ═══════════════════════════════════════════════════════════════════════════════

def open_book(book: dict) -> bool:
    """
    Open the book in ADE.  Tries known ADE executable paths first
    (avoids the Acrobat DRM error); falls back to os.startfile.
    """
    path: Path = book["path"]

    ade_candidates = [
        Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"))
            / "Adobe" / "Adobe Digital Editions 4.5" / "DigitalEditions.exe",
        Path(os.environ.get("ProgramFiles", r"C:\Program Files"))
            / "Adobe" / "Adobe Digital Editions 4.5" / "DigitalEditions.exe",
        Path(os.environ.get("LocalAppData", ""))
            / "Adobe" / "Adobe Digital Editions 4.5" / "DigitalEditions.exe",
    ]

    ade_exe = next((p for p in ade_candidates if p.exists()), None)

    try:
        if ade_exe:
            info(f"Launching ADE: {ade_exe.name}")
            subprocess.Popen([str(ade_exe), str(path)])
        else:
            warn("ADE executable not found in default locations — using OS default.")
            warn("If Acrobat opens instead, set ADE as default for .epub files.")
            os.startfile(str(path))
        return True
    except Exception as exc:
        error(f"Could not open '{path.name}': {exc}")
        return False


# ═══════════════════════════════════════════════════════════════════════════════
# 11. MAIN CAPTURE FUNCTION
# ═══════════════════════════════════════════════════════════════════════════════

def _check_capture_deps() -> list[str]:
    missing = []
    if not HAS_MSS:        missing.append("mss")
    if not HAS_PIL:        missing.append("Pillow")
    if not HAS_REPORTLAB:  missing.append("reportlab")
    return missing


def capture_book_to_pdf(
    book:       dict,
    output_dir: Optional[Path],
    page_delay: float = 1.0,
    max_pages:  int   = 500,
    open_wait:  float = 10.0,
    quality:    int   = 85,
    dpi:        int   = 150,
    crop_w:     int   = DEFAULT_CROP_W,
    crop_h:     int   = DEFAULT_CROP_H,
    crop_top:   int   = DEFAULT_TOP,
    crop_left:  int   = DEFAULT_LEFT,
) -> Optional[Path]:
    """
    Capture every page of *book* as displayed in ADE, apply a fixed offset
    crop to remove the ADE chrome, JPEG-compress, and assemble a single PDF.

    Crop layout (3840×2160 default)
    ────────────────────────────────
      top    = 125 px  — ADE toolbar + menu row
      left   = 516 px  — grey left margin
      width  = 2794 px — page content
      height = 1962 px — page content
      bottom = 125 + 1962 = 2087 → bottom margin = 73 px
      right  = 516 + 2794 = 3310 → right  margin = 530 px

    First-page wait
    ───────────────
    Uses _wait_for_stable_page() — polls until content hash stops changing,
    ensuring the page is fully rendered before capture.
    """
    missing = _check_capture_deps()
    if missing:
        error(f"Screen capture requires: {', '.join(missing)}")
        info(f"Install:  pip install {' '.join(missing)}")
        return None

    title:   str   = book["title"]
    out_dir: Path  = output_dir or book["path"].parent

    safe_title = "".join(c if c.isalnum() or c in " _-" else "_" for c in title).strip()
    timestamp  = datetime.now().strftime("%Y%m%d_%H%M%S")
    pdf_path   = out_dir / f"{safe_title}_{timestamp}_capture.pdf"

    header(f"📸  Capturing: {title}")
    info(f"Output PDF   : {pdf_path}")
    info(f"Page delay   : {page_delay}s  |  Max pages: {max_pages}")
    info(f"JPEG quality : {quality}  |  PDF DPI: {dpi}")
    info(f"Crop offset  : top={crop_top}px  left={crop_left}px  →  {crop_w}×{crop_h}px")

    # ── Open book ─────────────────────────────────────────────────────────────
    if not open_book(book):
        return None

    info(f"Waiting {open_wait}s for ADE to launch and load the book…")
    time.sleep(open_wait)

    # ── Find ADE window ───────────────────────────────────────────────────────
    hwnd = wait_for_ade_hwnd(timeout=30.0)
    if hwnd is None:
        error("ADE window not found. Is Adobe Digital Editions installed?")
        return None

    title_bar = _get_window_title(hwnd)
    success(f"Found ADE window: '{title_bar}'  (HWND={hwnd})")

    focus_and_maximise(hwnd)

    # ── Wait for page 1 to FULLY render ───────────────────────────────────────
    info("Waiting for first page to fully render (stable hash)…")
    first_raw = _wait_for_stable_page(hwnd, timeout=STABLE_TIMEOUT,
                                      poll=STABLE_POLL, streak=STABLE_HASH_COUNT)
    if first_raw is None or _is_black(first_raw):
        error("First page is black or did not render.")
        error("Make sure ADE has fully opened the book, then try again.")
        return None

    b = _brightness(first_raw)
    W, H = first_raw.size
    success(f"First page rendered  (brightness {b:.0f}/255,  full size {W}×{H}px)")

    # ── Apply fixed offset crop to page 1 ────────────────────────────────────
    first_img = offset_crop(first_raw, crop_top, crop_left, crop_w, crop_h)
    bottom_margin = H - crop_top - crop_h
    success(f"Page 1 cropped: top={crop_top}px  left={crop_left}px  "
            f"bottom={bottom_margin}px  →  {first_img.width}×{first_img.height}px")

    # ── Capture loop ──────────────────────────────────────────────────────────
    jpeg_frames: list[bytes] = [_compress_jpeg(first_img, quality)]
    captured     = 1
    prev_hash    = _img_hash(first_img)
    dup_streak   = 0
    black_streak = 0
    MAX_BLACK    = 5

    sys.stdout.write(_color(
        f"  📄  Page {captured:>4} captured  ({len(jpeg_frames[-1])//1024}KB)\r",
        Fore.WHITE
    ))
    sys.stdout.flush()

    info("\nCapture running — you can use your computer normally.")
    info("Ctrl+C to stop early and save what has been captured.\n")

    try:
        for _ in range(2, max_pages + 1):

            # Post Right-Arrow directly to ADE's message queue — no focus needed
            send_next_page(hwnd)
            time.sleep(page_delay)

            # Wait for the new page to paint
            raw_img = _wait_for_render(hwnd)
            if raw_img is None:
                warn("Cannot capture ADE window — stopping.")
                break

            # ── Black frame guard ─────────────────────────────────────────────
            if _is_black(raw_img):
                black_streak += 1
                warn(f"Black frame #{black_streak} "
                     f"(brightness {_brightness(raw_img):.0f}) — skipping.")
                if black_streak >= MAX_BLACK:
                    warn("Too many consecutive black frames — stopping.")
                    break
                continue
            black_streak = 0

            # ── Apply fixed offset crop to remove chrome ──────────────────────
            img = offset_crop(raw_img, crop_top, crop_left, crop_w, crop_h)

            # ── End-of-book detection ─────────────────────────────────────────
            h = _img_hash(img)
            if h == prev_hash:
                dup_streak += 1
                if dup_streak >= END_OF_BOOK_STREAK:
                    success(f"End of book detected after {captured} page(s).")
                    break
                continue
            dup_streak = 0
            prev_hash  = h

            # ── Compress and store ────────────────────────────────────────────
            jpeg = _compress_jpeg(img, quality)
            jpeg_frames.append(jpeg)
            captured += 1

            sys.stdout.write(_color(
                f"  📄  Page {captured:>4} captured  ({len(jpeg)//1024}KB)\r",
                Fore.WHITE
            ))
            sys.stdout.flush()

    except KeyboardInterrupt:
        warn("\nCapture stopped by user (Ctrl+C).")

    print()

    if not jpeg_frames:
        error("No pages captured.")
        return None

    success(f"Captured {len(jpeg_frames)} page(s).")
    total_kb = sum(len(j) for j in jpeg_frames) // 1024
    info(f"Total JPEG data: {total_kb} KB → building PDF…")

    try:
        out_dir.mkdir(parents=True, exist_ok=True)
        _build_pdf(jpeg_frames, pdf_path, dpi=dpi)
        pdf_kb = pdf_path.stat().st_size // 1024
        success(f"PDF saved → {pdf_path}  ({pdf_kb} KB, {len(jpeg_frames)} pages)")
        return pdf_path
    except Exception as exc:
        error(f"PDF assembly failed: {exc}")
        return None


# ═══════════════════════════════════════════════════════════════════════════════
# 12. DISPLAY HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def print_book_row(idx: int, book: dict) -> None:
    title  = book["title"][:42].ljust(43)
    author = book["author"][:25].ljust(26)
    pages  = book.get("pages", 0)
    pg_str = f"{pages}pp" if pages else "?pp"
    print(f"  {str(idx).rjust(3)}.  "
          f"{_color(title, Fore.WHITE + Style.BRIGHT)}  "
          f"{_color(author, Fore.CYAN)}  "
          f"[{book['format']}] {pg_str}  {book['size_mb']:.1f}MB")

def print_book_detail(book: dict) -> None:
    header(f"📖  {book['title']}")
    pages     = book.get("pages", 0)
    pages_str = f"{pages}" if pages else "Unknown"
    if book.get("pages_note"):
        pages_str += f"  ({book['pages_note']})"
    for label, value in [
        ("Author",    book["author"]),
        ("Publisher", book["publisher"]),
        ("Date",      book["date"]),
        ("Language",  book["language"]),
        ("Format",    book["format"]),
        ("Pages",     pages_str),
        ("File size", f"{book['size_mb']} MB"),
        ("Modified",  book["modified"]),
        ("Path",      str(book["path"])),
    ]:
        print(f"  {_color(label.ljust(12), Fore.YELLOW)}  {value}")


# ═══════════════════════════════════════════════════════════════════════════════
# 13. INTERACTIVE MENUS
# ═══════════════════════════════════════════════════════════════════════════════

def menu_shelf(library: ADELibrary, cfg: dict) -> None:
    names = library.shelf_names()
    if not names:
        warn("No bookshelves found.")
        return
    header("📚  Your Bookshelves")
    for i, name in enumerate(names, 1):
        count = len(library.shelves[name])
        print(f"  {str(i).rjust(3)}.  {_color(name, Fore.GREEN)}  "
              f"({count} book{'s' if count != 1 else ''})")
    choice = prompt("Shelf number (ENTER = back)")
    if not choice:
        return
    try:
        idx = int(choice) - 1
        if 0 <= idx < len(names):
            menu_books(library, names[idx], cfg)
        else:
            warn("Invalid number.")
    except ValueError:
        warn("Please enter a number.")


def menu_books(library: ADELibrary, shelf_name: str, cfg: dict) -> None:
    books = library.shelves.get(shelf_name, [])
    if not books:
        warn(f"Shelf '{shelf_name}' is empty.")
        return
    while True:
        header(f"📖  {shelf_name}  ({len(books)} books)")
        for i, book in enumerate(books, 1):
            print_book_row(i, book)
        print(_color("\n  [number] view/open/capture   [b] back   [q] quit", Fore.YELLOW))
        choice = prompt("Select")
        if choice.lower() == "b":
            return
        if choice.lower() == "q":
            sys.exit(0)
        try:
            idx = int(choice) - 1
            if 0 <= idx < len(books):
                menu_book_detail(books[idx], cfg)
            else:
                warn("Invalid number.")
        except ValueError:
            warn("Enter a number, 'b', or 'q'.")


def menu_book_detail(book: dict, cfg: dict) -> None:
    print_book_detail(book)
    print()
    print(_color("  [o]  Open in Adobe Digital Editions", Fore.WHITE))
    print(_color("  [c]  Capture all pages → compressed PDF", Fore.WHITE))
    print(_color("  [b]  Back", Fore.WHITE))
    choice = prompt("Action").lower()
    if choice == "o":
        if open_book(book):
            success("Book launched in ADE.")
        else:
            error("Failed to open.")
    elif choice == "c":
        _run_capture_flow(book, cfg)


def _run_capture_flow(book: dict, cfg: dict) -> None:
    missing = _check_capture_deps()
    if missing:
        error(f"Cannot capture — missing packages: {', '.join(missing)}")
        info(f"Install:  pip install {' '.join(missing)}")
        return

    # Use the book's own detected page count as the default max, if available
    book_pages   = book.get("pages", 0)
    default_max  = book_pages if book_pages > 0 else cfg["max_pages"]
    pages_source = f"from book metadata" if book_pages > 0 else "global default"

    info("Capture settings:")
    info(f"  Page delay   : {cfg['delay']}s")
    info(f"  Max pages    : {cfg['max_pages']}")
    info(f"  JPEG quality : {cfg['quality']}  (1=smallest, 95=best)")
    info(f"  Top offset   : {cfg['crop_top']}px")
    info(f"  Left offset  : {cfg['crop_left']}px")
    info(f"  Crop size    : {cfg['crop_w']}×{cfg['crop_h']}px")
    info(f"  Bottom margin: {2160 - cfg['crop_top'] - cfg['crop_h']}px  (calculated)")
    info(f"  Output dir   : {cfg['output'] or 'same folder as book'}")
    print()

    raw = prompt(f"Page delay [{cfg['delay']}s] (ENTER = keep)")
    try:    delay = float(raw) if raw else cfg["delay"]
    except: delay = cfg["delay"]

    raw2 = prompt(f"JPEG quality [{cfg['quality']}] (ENTER = keep)")
    try:    quality = int(raw2) if raw2 else cfg["quality"]
    except: quality = cfg["quality"]

    raw3 = prompt(f"Top offset [{cfg['crop_top']}px] (ENTER = keep)")
    try:    crop_top = int(raw3) if raw3 else cfg["crop_top"]
    except: crop_top = cfg["crop_top"]

    raw4 = prompt(f"Left offset [{cfg['crop_left']}px] (ENTER = keep)")
    try:    crop_left = int(raw4) if raw4 else cfg["crop_left"]
    except: crop_left = cfg["crop_left"]

    raw5 = prompt(f"Crop width  [{cfg['crop_w']}px] (ENTER = keep)")
    try:    crop_w = int(raw5) if raw5 else cfg["crop_w"]
    except: crop_w = cfg["crop_w"]

    raw6 = prompt(f"Crop height [{cfg['crop_h']}px] (ENTER = keep)")
    try:    crop_h = int(raw6) if raw6 else cfg["crop_h"]
    except: crop_h = cfg["crop_h"]

    raw7 = prompt(f"Max pages [{cfg['max_pages']}] (ENTER = keep)")
    try:    max_pages = int(raw7) if raw7 else cfg["max_pages"]
    except: max_pages = cfg["max_pages"]

    print()
    warn("ADE will open the book and start capturing automatically.")
    warn("You can use your computer normally during capture.")
    warn("Ctrl+C in this terminal stops capture early.")

    if prompt("Start capture? [y/n]").lower() != "y":
        info("Cancelled.")
        return

    pdf_path = capture_book_to_pdf(
        book       = book,
        output_dir = cfg["output"],
        page_delay = delay,
        max_pages  = max_pages,
        quality    = quality,
        dpi        = cfg["dpi"],
        crop_w     = crop_w,
        crop_h     = crop_h,
        crop_top   = crop_top,
        crop_left  = crop_left,
    )
    if pdf_path:
        success(f"Done!  PDF:\n  {pdf_path}")
    else:
        error("Capture failed — see messages above.")


def menu_search(library: ADELibrary, cfg: dict) -> None:
    q = prompt("Search by title or author")
    if not q:
        return
    results = library.search(q)
    if not results:
        warn(f"No books found for '{q}'.")
        return
    header(f"🔍  Results for '{q}'  ({len(results)} found)")
    for i, book in enumerate(results, 1):
        print_book_row(i, book)
    choice = prompt("Open a book (number) or ENTER to go back")
    if not choice:
        return
    try:
        idx = int(choice) - 1
        if 0 <= idx < len(results):
            menu_book_detail(results[idx], cfg)
        else:
            warn("Invalid number.")
    except ValueError:
        warn("Please enter a number.")


def menu_all_books(library: ADELibrary, cfg: dict) -> None:
    if not library.books:
        warn("No books found.")
        return
    books = sorted(library.books, key=lambda b: b["title"].lower())
    header(f"📚  All Books  ({len(books)} total)")
    for i, book in enumerate(books, 1):
        print_book_row(i, book)
    choice = prompt("Open a book (number) or ENTER to go back")
    if not choice:
        return
    try:
        idx = int(choice) - 1
        if 0 <= idx < len(books):
            menu_book_detail(books[idx], cfg)
        else:
            warn("Invalid number.")
    except ValueError:
        warn("Please enter a number.")


def menu_stats(library: ADELibrary, cfg: dict) -> None:
    total    = len(library.books)
    epubs    = sum(1 for b in library.books if b["format"] == "EPUB")
    pdfs     = sum(1 for b in library.books if b["format"] == "PDF")
    total_mb = sum(b["size_mb"] for b in library.books)
    header("📊  Library Statistics")
    for label, value in [
        ("Total books",   total),
        ("EPUB files",    epubs),
        ("PDF files",     pdfs),
                ("Bookshelves",   len(library.shelves)),
        ("Total size",    f"{total_mb:.1f} MB"),
        ("Library path",  str(library.root)),
        ("JPEG quality",  cfg["quality"]),
        ("PDF DPI",       cfg["dpi"]),
        ("Top offset",    f"{cfg['crop_top']}px"),
        ("Left offset",   f"{cfg['crop_left']}px"),
        ("Crop size",     f"{cfg['crop_w']}×{cfg['crop_h']}px"),
        ("Bottom margin", f"{2160 - cfg['crop_top'] - cfg['crop_h']}px  (calculated)"),
    ]:
        print(f"  {_color(label.ljust(16), Fore.YELLOW)}  {value}")
    missing = _check_capture_deps()
    print()
    if missing:
        warn(f"Screen capture unavailable (missing: {', '.join(missing)})")
        info(f"Install:  pip install {' '.join(missing)}")
    else:
        success("Windows native capture (mss) is ready ✔")


# ═══════════════════════════════════════════════════════════════════════════════
# 14. MAIN LOOP
# ═══════════════════════════════════════════════════════════════════════════════

MAIN_MENU = """
  ┌─────────────────────────────────────────┐
  │  [1]  Browse Bookshelves                │
  │  [2]  All Books                         │
  │  [3]  Search Library                    │
  │  [4]  Library Statistics                │
  │  [q]  Quit                              │
  └─────────────────────────────────────────┘"""


def run(library_path: Optional[Path], cfg: dict) -> None:
    if library_path is None:
        library_path = detect_ade_library()

    if library_path is None or not library_path.exists():
        error("ADE library folder not found.")
        info('Pass the path with:  python run_ade.py --library "C:\\...\\My Digital Editions"')
        sys.exit(1)

    header("Adobe Digital Editions — Bookshelf Manager + Windows Capture")
    info(f"Library : {library_path}")
    info("Scanning…")

    library = ADELibrary(library_path)
    if not library.books:
        warn("No EPUB or PDF files found — has ADE downloaded any books?")
        sys.exit(0)

    success(f"Found {len(library.books)} book(s) on {len(library.shelves)} shelf/shelves.")

    missing = _check_capture_deps()
    if missing:
        warn(f"Screen capture not available (missing: {', '.join(missing)})")
        info(f"Install:  pip install {' '.join(missing)}")
    else:
        success("Windows native capture ready — select a book and press [c].")

    while True:
        print(_color(MAIN_MENU, Fore.CYAN))
        choice = prompt("Choose an option")
        if   choice == "1": menu_shelf(library, cfg)
        elif choice == "2": menu_all_books(library, cfg)
        elif choice == "3": menu_search(library, cfg)
        elif choice == "4": menu_stats(library, cfg)
        elif choice.lower() == "q":
            success("Goodbye!")
            break
        else:
            warn("Choose 1–4 or q.")


# ═══════════════════════════════════════════════════════════════════════════════
# 15. ARGUMENT PARSING & ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════════

def parse_args():
    p = argparse.ArgumentParser(
        description="ADE Bookshelf Manager — Windows Native Capture → Compressed PDF",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Recommended usage (handles venv + path quoting automatically):
  python run_ade.py
  python run_ade.py --library "C:\\Users\\HP Z4 G4 Workstation\\Documents\\My Digital Editions"
  python run_ade.py --quality 75 --dpi 120 --delay 1.5
        """
    )
    p.add_argument("--library",   "-l", metavar="PATH", default=None)
    p.add_argument("--output",    "-o", metavar="DIR",  default=None)
    p.add_argument("--delay",     "-d", metavar="SECS", type=float, default=1.0)
    p.add_argument("--max-pages", "-m", metavar="N",    type=int,   default=500)
    p.add_argument("--quality",   "-q", metavar="N",    type=int,   default=85,
                   help="JPEG quality 1-95 (default 85)")
    p.add_argument("--dpi",             metavar="N",    type=int,   default=150,
                   help="PDF resolution hint in DPI (default 150)")
    p.add_argument("--crop-w",          metavar="PX",   type=int,   default=DEFAULT_CROP_W,
                   help=f"Crop width in pixels (default {DEFAULT_CROP_W})")
    p.add_argument("--crop-h",          metavar="PX",   type=int,   default=DEFAULT_CROP_H,
                   help=f"Crop height in pixels (default {DEFAULT_CROP_H})")
    p.add_argument("--top",             metavar="PX",   type=int,   default=DEFAULT_TOP,
                   help=f"Top offset in pixels (default {DEFAULT_TOP})")
    p.add_argument("--left",            metavar="PX",   type=int,   default=DEFAULT_LEFT,
                   help=f"Left offset in pixels (default {DEFAULT_LEFT})")
    return p.parse_args()


if __name__ == "__main__":
    args = parse_args()
    cfg  = {
        "delay":     args.delay,
        "max_pages": args.max_pages,
        "quality":   max(1, min(95, args.quality)),
        "dpi":       args.dpi,
        "crop_w":    args.crop_w,
        "crop_h":    args.crop_h,
        "crop_top":  args.top,
        "crop_left": args.left,
        "output":    normalize_path(args.output) if args.output else None,
    }
    lib_path = normalize_path(args.library) if args.library else None
    run(lib_path, cfg)
