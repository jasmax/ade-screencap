#!/usr/bin/env python3

import os
import sys
import shutil
import platform
import subprocess
import argparse
from pathlib import Path

# ── Constants ─────────────────────────────────────────────────────────────────

SCRIPT_DIR   = Path(__file__).resolve().parent
MAIN_SCRIPT  = SCRIPT_DIR / "ade_bookshelf.py"
DEFAULT_VENV = SCRIPT_DIR / ".venv"

REQUIRED_PACKAGES = [
    "colorama",
    "mss",       # Windows native DXGI screen capture
    "Pillow",    # image processing, JPEG compression, PDF assembly
    "reportlab", # PDF fallback assembler
]

IS_WINDOWS = platform.system() == "Windows"


# ── Console helpers ───────────────────────────────────────────────────────────

def _tag(tag: str, msg: str) -> None:
    print(f"  [{tag}] {msg}", flush=True)

def ok(msg: str)   -> None: _tag("✔", msg)
def info(msg: str) -> None: _tag("ℹ", msg)
def warn(msg: str) -> None: _tag("⚠", msg)
def err(msg: str)  -> None: _tag("✖", msg)
def sep()          -> None: print()


# ── Path helpers ──────────────────────────────────────────────────────────────

def venv_python(venv: Path) -> Path:
    return venv / ("Scripts\\python.exe" if IS_WINDOWS else "bin/python")

def venv_pip(venv: Path) -> Path:
    return venv / ("Scripts\\pip.exe" if IS_WINDOWS else "bin/pip")

def normalize_path(raw: str) -> Path:
    return Path(raw.strip().strip("\"'")).expanduser().resolve()


# ── Venv helpers ──────────────────────────────────────────────────────────────

def is_venv_healthy(venv: Path) -> bool:
    py = venv_python(venv)
    return py.exists() and os.access(str(py), os.X_OK)

def running_inside_venv() -> bool:
    return (
        hasattr(sys, "real_prefix")
        or (hasattr(sys, "base_prefix") and sys.base_prefix != sys.prefix)
    )

def create_venv(venv: Path) -> None:
    info(f"Creating virtual environment at:\n       {venv}")
    try:
        subprocess.run([sys.executable, "-m", "venv", str(venv)], check=True)
        ok("Virtual environment created.")
    except subprocess.CalledProcessError as exc:
        err(f"Failed to create venv: {exc}")
        sys.exit(1)

def upgrade_pip(venv: Path) -> None:
    subprocess.run(
        [str(venv_python(venv)), "-m", "pip", "install", "--upgrade", "pip"],
        check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )

def install_packages(venv: Path, packages: list, upgrade: bool = False) -> None:
    info(f"Installing: {', '.join(packages)}")
    cmd = [str(venv_pip(venv)), "install"] + (["--upgrade"] if upgrade else []) + packages
    try:
        subprocess.run(cmd, check=True)
        ok("Packages installed.")
    except subprocess.CalledProcessError as exc:
        err(f"pip install failed: {exc}")
        warn("Check your internet connection and try again.")
        sys.exit(1)

def packages_installed(venv: Path) -> bool:
    import_map = {
        "colorama":  "colorama",
        "mss":       "mss",
        "Pillow":    "PIL",
        "reportlab": "reportlab",
    }
    py = str(venv_python(venv))
    for pkg, mod in import_map.items():
        r = subprocess.run(
            [py, "-c", f"import {mod}"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        if r.returncode != 0:
            info(f"Package not found in venv: {pkg}")
            return False
    return True


# ── Bootstrap ─────────────────────────────────────────────────────────────────

def bootstrap(venv: Path, refresh: bool) -> None:
    if refresh and venv.exists():
        warn("--refresh: removing existing venv…")
        shutil.rmtree(venv)

    if not is_venv_healthy(venv):
        create_venv(venv)
        upgrade_pip(venv)
        install_packages(venv, REQUIRED_PACKAGES)
    else:
        ok(f"Venv OK: {venv}")
        if refresh or not packages_installed(venv):
            install_packages(venv, REQUIRED_PACKAGES, upgrade=refresh)
        else:
            ok("All packages already installed.")


# ── Launch ────────────────────────────────────────────────────────────────────

def launch(venv: Path, forward_args: list) -> None:
    """
    Execute ade_bookshelf.py inside the venv Python.
    Each path is a separate list element — Python handles OS quoting
    internally so spaces in usernames/folders are never misinterpreted.
    """
    if not MAIN_SCRIPT.exists():
        err(f"Cannot find ade_bookshelf.py at: {MAIN_SCRIPT}")
        err("Both files must be in the same folder.")
        sys.exit(1)

    py  = str(venv_python(venv))
    cmd = [py, str(MAIN_SCRIPT)] + forward_args

    info(f"Python : {py}")
    info(f"Script : {MAIN_SCRIPT}")
    sep()

    try:
        result = subprocess.run(cmd)
        sys.exit(result.returncode)
    except KeyboardInterrupt:
        sys.exit(0)


# ── Forward arg builder ───────────────────────────────────────────────────────

def build_forward_args(known, extra: list, library_raw: str) -> list:
    """
    Build the argv list for ade_bookshelf.py.
    Every value is a separate list element — no shell-string concatenation.
    """
    forward = []

    if library_raw:
        lib_path = normalize_path(library_raw)
        info(f"Library path : {lib_path}")
        if not lib_path.exists():
            warn(f"Path does not exist: {lib_path}")
        forward += ["--library", str(lib_path)]

    if known.output:
        forward += ["--output", str(normalize_path(known.output))]

    if known.delay is not None:
        forward += ["--delay", str(known.delay)]

    if known.max_pages is not None:
        forward += ["--max-pages", str(known.max_pages)]

    if known.quality is not None:
        forward += ["--quality", str(known.quality)]

    if known.dpi is not None:
        forward += ["--dpi", str(known.dpi)]

    if known.crop_w is not None:
        forward += ["--crop-w", str(known.crop_w)]

    if known.crop_h is not None:
        forward += ["--crop-h", str(known.crop_h)]

    if known.top is not None:
        forward += ["--top", str(known.top)]

    if known.left is not None:
        forward += ["--left", str(known.left)]

    forward += extra
    return forward


# ── Interactive path prompt ───────────────────────────────────────────────────

def prompt_library_path() -> str:
    print()
    print("  ┌──────────────────────────────────────────────────────────┐")
    print("  │  No --library path provided.                             │")
    print("  │  Paste your ADE library path below and press Enter.      │")
    print("  │  No quotes needed — paste the path exactly as-is.        │")
    print("  │                                                           │")
    print("  │  Default:                                                 │")
    print("  │  C:\\Users\\<YourName>\\Documents\\My Digital Editions        │")
    print("  └──────────────────────────────────────────────────────────┘")
    print()
    return input("  Library path (or press Enter to auto-detect): ").strip()


# ── Argument parsing ──────────────────────────────────────────────────────────

def parse_args():
    p = argparse.ArgumentParser(
        add_help=False,
        description="Venv launcher for ade_bookshelf.py",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("--venv",      metavar="PATH", default=None)
    p.add_argument("--refresh",   action="store_true")
    p.add_argument("--library",   "-l", metavar="PATH",  default=None)
    p.add_argument("--output",    "-o", metavar="DIR",   default=None)
    p.add_argument("--delay",     "-d", metavar="SECS",  type=float, default=None)
    p.add_argument("--max-pages", "-m", metavar="N",     type=int,   default=None)
    p.add_argument("--quality",   "-q", metavar="N",     type=int,   default=None)
    p.add_argument("--dpi",             metavar="N",     type=int,   default=None)
    p.add_argument("--crop-w",          metavar="PX",    type=int,   default=None,
                   help="Crop width px (default 2794)")
    p.add_argument("--crop-h",          metavar="PX",    type=int,   default=None,
                   help="Crop height px (default 1962)")
    p.add_argument("--top",             metavar="PX",    type=int,   default=None,
                   help="Top offset px (default 127)")
    p.add_argument("--left",            metavar="PX",    type=int,   default=None,
                   help="Left offset px (default 523)")
    known, extra = p.parse_known_args()
    return known, extra


# ── Entry point ───────────────────────────────────────────────────────────────

def main() -> None:
    print()
    print("  ══════════════════════════════════════════════════════════")
    print("   ADE Bookshelf Manager — Windows Native Capture Launcher")
    print("  ══════════════════════════════════════════════════════════")
    sep()

    known, extra = parse_args()
    venv = Path(known.venv) if known.venv else DEFAULT_VENV

    if running_inside_venv() and not known.refresh:
        info("Already inside a virtual environment — skipping bootstrap.")
        library_raw = known.library or ""
        launch(venv, build_forward_args(known, extra, library_raw))
        return

    info(f"Python   : {sys.executable}  (v{platform.python_version()})")
    info(f"Platform : {platform.system()} {platform.machine()}")
    info(f"Venv     : {venv}")
    sep()

    if not IS_WINDOWS:
        err("This tool uses Windows-native capture APIs.")
        err("It must run on Windows.")
        sys.exit(1)

    library_raw = known.library or ""
    if not library_raw:
        library_raw = prompt_library_path()

    bootstrap(venv, refresh=known.refresh)
    sep()

    launch(venv, build_forward_args(known, extra, library_raw))


if __name__ == "__main__":
    main()
