ade_bookshelf.py
================
Adobe Digital Editions (ADE) Bookshelf Manager  +  Windows Screen Capture
--------------------------------------------------------------------------

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  RECOMMENDED: always launch via run_ade.py
      python run_ade.py
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

How Windows capture works (NEW)
────────────────────────────────
  • Uses  mss  (Windows native DXGI / BitBlt screen capture) instead of
    pyautogui.  mss calls the Win32 GDI BitBlt API directly, which:
      - Captures GPU-composited frames correctly (no black screens)
      - Works even when the window is partially obscured
      - Is 5-10× faster than pyautogui's PIL-based capture
  • Uses ctypes + win32 (built-in) to find, restore, maximise, and
    foreground the ADE window — no pygetwindow dependency.
  • Each captured frame is JPEG-compressed (quality 85) before being
    embedded into the output PDF — dramatically smaller file sizes vs PNG.
  • PDF assembly uses Pillow's built-in PDF save (no reportlab needed for
    the simple case; reportlab used for multi-image layout control).

Capture flow
────────────
  1. Open book → ADE launches.
  2. Find ADE window via EnumWindows (Win32 API via ctypes).
  3. Restore + maximise + foreground ADE.
  4. Capture page 1 with mss, verify it is not black.
  5. Loop:
       a. Send WM_KEYDOWN / WM_KEYUP Right-Arrow to ADE's HWND directly
          (no pyautogui keyboard, no focus required).
       b. Poll mss until frame brightness > threshold (render guard).
       c. JPEG-compress and store frame in memory.
       d. Hash 64×64 thumbnail to detect end-of-book.
  6. Compress all JPEG frames into a single PDF via Pillow.

CLI options
───────────
  --library   PATH   ADE library folder (auto-detected if omitted)
  --output    DIR    Folder for captured PDFs (default: beside book file)
  --delay     SECS   Base wait between page turns (default 1.0)
  --max-pages N      Hard cap on captured pages (default 500)
  --quality   N      JPEG quality 1-95 (default 85; lower = smaller file)
  --dpi       N      Output PDF DPI hint (default 150)

Dependencies (installed automatically by run_ade.py)
─────────────────────────────────────────────────────
  colorama  mss  Pillow  reportlab
  (pygetwindow and pyautogui are NO LONGER required)

run_ade.py
==========
Self-bootstrapping launcher for ade_bookshelf.py
-------------------------------------------------
Run this ONE script with your system Python. It will:

  1. Create a virtual environment (.venv/) beside this file.
  2. Install all required packages into the venv.
  3. Re-launch ade_bookshelf.py inside the venv with all paths
     correctly quoted so spaces in folder/usernames never break anything.

Required packages (installed automatically)
───────────────────────────────────────────
  colorama   — coloured terminal output
  mss        — Windows native DXGI/BitBlt screen capture
  Pillow     — image processing + JPEG compression + PDF assembly
  reportlab  — PDF fallback assembler

  NOTE: pyautogui and pygetwindow are NO LONGER needed.

Usage (PowerShell / VS Code terminal)
──────────────────────────────────────
  python run_ade.py
  python run_ade.py --library "path to library"
  python run_ade.py --quality 75 --dpi 120 --delay 1.5
  python run_ade.py --refresh        # force reinstall of all packages
  python run_ade.py --venv PATH      # custom venv location
