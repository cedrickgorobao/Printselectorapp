"""
auto_extract.py
Extracts pages from Publisher files for each variety (Arabica, Barako, Robusta, Pahimis Special),
saves them to pages/<variety>/ folders, then starts a local web server and opens the app.

First run: installs pywin32 automatically if not present.

─────────────────────────────────────────────────────────
CONFIGURE: Set the path to each .pub file below.
           Leave as None if a variety has no file yet.
─────────────────────────────────────────────────────────
"""

import os
import sys
import subprocess
import time
import threading
import webbrowser
from pathlib import Path

# ── CONFIGURE YOUR .pub FILES HERE ────────────────────────────────────────────
PUB_FILES = {
    "arabica": r"C:\Users\I508865\Downloads\Arabica.pub",
    "barako":  r"C:\Users\I508865\Downloads\Barako.pub",
    "robusta": r"C:\Users\I508865\Downloads\Robusta.pub",
    "pahimis": r"C:\Users\I508865\Downloads\Pahimis Special.pub",
}

# Labels for display
LABELS = {
    "arabica": "Arabica",
    "barako":  "Barako",
    "robusta": "Robusta",
    "pahimis": "Pahimis Special",
}

SCRIPT_DIR = Path(__file__).parent.resolve()
PAGES_DIR  = SCRIPT_DIR / "pages"
PORT       = 8765
PNG_FORMAT = 3   # Publisher constant: pbExportFormatPNG


# ── ENSURE PYWIN32 ────────────────────────────────────────────────────────────
def ensure_pywin32():
    try:
        import win32com.client  # noqa
        return
    except ImportError:
        pass

    print("=" * 55)
    print("  pywin32 not found — installing now (one-time setup)...")
    print("=" * 55)
    result = subprocess.run(
        [sys.executable, "-m", "pip", "install", "pywin32"]
    )
    if result.returncode != 0:
        print("\nERROR: Failed to install pywin32.")
        print("Please run manually:  pip install pywin32")
        input("\nPress Enter to exit...")
        sys.exit(1)

    # Register DLLs
    scripts_dir  = Path(sys.executable).parent / "Scripts"
    postinstall  = scripts_dir / "pywin32_postinstall.py"
    if postinstall.exists():
        subprocess.run([sys.executable, str(postinstall), "-install"],
                       capture_output=True)

    try:
        import win32com.client  # noqa
        print("pywin32 installed OK.\n")
    except ImportError:
        print("\nInstall succeeded but import still fails.")
        print("Please close this window and run AUTO_EXTRACT.bat again.")
        input("Press Enter to exit...")
        sys.exit(1)


# ── EXTRACT PAGES FOR ONE VARIETY ─────────────────────────────────────────────
def extract_variety(slug, pub_path):
    label     = LABELS[slug]
    out_dir   = PAGES_DIR / slug
    pub_path  = Path(pub_path)

    if not pub_path.exists():
        print(f"  [{label}] WARNING: File not found — skipping")
        print(f"             Expected: {pub_path}")
        return

    out_dir.mkdir(parents=True, exist_ok=True)
    existing = list(out_dir.glob("page*.png"))

    if len(existing) >= 12:
        print(f"  [{label}] Already extracted ({len(existing)} files). Skipping.")
        return

    import win32com.client

    print(f"\n  [{label}] Opening: {pub_path.name}")
    app = win32com.client.Dispatch("Publisher.Application")
    app.Visible = True

    doc   = app.Open(str(pub_path))
    total = doc.Pages.Count
    print(f"  [{label}] {total} pages found. Exporting...")

    for i in range(1, total + 1):
        out = str(out_dir / f"page{i:02d}.png")
        print(f"    Page {i:2d}/{total}", end="\r", flush=True)
        doc.Pages.Item(i).Export(out, PNG_FORMAT)

    doc.Close(False)
    app.Quit()
    print(f"  [{label}] Done — {total} pages saved to pages/{slug}/")


# ── EXTRACT ALL ───────────────────────────────────────────────────────────────
def extract_all():
    PAGES_DIR.mkdir(parents=True, exist_ok=True)

    print("\n" + "=" * 55)
    print("  Extracting pages from Publisher files...")
    print("=" * 55)

    for slug, path in PUB_FILES.items():
        if path is None:
            print(f"  [{LABELS[slug]}] No file configured — skipping")
            continue
        try:
            extract_variety(slug, path)
        except Exception as e:
            print(f"  [{LABELS[slug]}] ERROR: {e}")

    print("\nExtraction complete!\n")


# ── LOCAL SERVER ──────────────────────────────────────────────────────────────
def start_server():
    from http.server import HTTPServer, SimpleHTTPRequestHandler

    class Silent(SimpleHTTPRequestHandler):
        def log_message(self, fmt, *args):
            pass

    os.chdir(SCRIPT_DIR)
    httpd = HTTPServer(("localhost", PORT), Silent)
    httpd.serve_forever()


# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    print()
    print("=" * 55)
    print("  XMAS Page Printer — Auto Extract & Launch")
    print("=" * 55)
    print()
    print("  Varieties: Arabica | Barako | Robusta | Pahimis Special")
    print()

    ensure_pywin32()
    extract_all()

    print("Starting local server...")
    t = threading.Thread(target=start_server, daemon=True)
    t.start()
    time.sleep(0.6)

    url = f"http://localhost:{PORT}"
    print(f"Opening browser: {url}\n")
    webbrowser.open(url)

    print("App is running! Keep this window open while using the printer app.")
    print("Press Ctrl+C to stop.\n")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\nServer stopped.")


if __name__ == "__main__":
    main()
