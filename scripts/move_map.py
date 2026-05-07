import os
import re
import shutil
import glob
import time
import traceback

# ── PATHS ─────────────────────────────────────────────────────────────
DOWNLOADS_FOLDER  = r"C:\Users\username\Downloads"
maps_FOLDER       = r"C:\Automation\maps"
TEMP_ADDRESS_FILE = r"C:\Automation\temp_mapaddress.txt"   # written by step 34 in flow
LOG_FILE          = r"C:\Automation\move_map_log.txt"

def log(msg):
    print(msg)
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(msg + '\n')

# Clear log at start
with open(LOG_FILE, 'w', encoding='utf-8') as f:
    f.write(f"=== move_map.py run at {time.strftime('%Y-%m-%d %H:%M:%S')} ===\n")

# ── ADDRESS NORMALIZER ────────────────────────────────────────────────
def norm_addr(addr):
    if not addr: return ''
    addr = str(addr).upper().strip()
    addr = re.sub(r'[.,]', '', addr)
    return re.sub(r'\s+', ' ', addr).strip()

# ── READ ADDRESS — try multiple encodings ─────────────────────────────
log(f"Reading address from {TEMP_ADDRESS_FILE}")
raw_address = None
for enc in ('utf-16', 'utf-8-sig', 'utf-8', 'latin-1'):
    try:
        with open(TEMP_ADDRESS_FILE, 'r', encoding=enc) as f:
            raw_address = f.read().strip()
        log(f"Read OK with encoding={enc}, address='{raw_address}'")
        break
    except Exception as e:
        log(f"  encoding={enc} failed: {e}")

if not raw_address:
    log("ERROR: Could not read address from temp_mapaddress.txt")
    raise SystemExit(1)

normalized_address = norm_addr(raw_address)
log(f"Normalized address = '{normalized_address}'")

# ── MAKE SURE maps FOLDER EXISTS ──────────────────────────────────────
os.makedirs(maps_FOLDER, exist_ok=True)

# ── FIND NEWEST FLOORPLAN PDF IN DOWNLOADS ────────────────────────────
pdf_files = (
    glob.glob(os.path.join(DOWNLOADS_FOLDER, "*loorplan*.pdf")) +
    glob.glob(os.path.join(DOWNLOADS_FOLDER, "*loorplan*.PDF"))
)

log(f"Found {len(pdf_files)} Floorplan PDF(s) in Downloads")

if not pdf_files:
    all_pdfs = glob.glob(os.path.join(DOWNLOADS_FOLDER, "*.pdf")) + \
               glob.glob(os.path.join(DOWNLOADS_FOLDER, "*.PDF"))
    log(f"No Floorplan PDFs found. Total PDFs in Downloads: {len(all_pdfs)}")
    if all_pdfs:
        all_pdfs.sort(key=os.path.getmtime, reverse=True)
        log(f"5 newest: {[os.path.basename(f) for f in all_pdfs[:5]]}")
    raise SystemExit(1)

pdf_files.sort(key=os.path.getmtime, reverse=True)
newest_pdf = pdf_files[0]
file_age_seconds = time.time() - os.path.getmtime(newest_pdf)

log(f"Newest Floorplan PDF = {os.path.basename(newest_pdf)}")
log(f"File age             = {int(file_age_seconds)} seconds")

if file_age_seconds > 600:
    log(f"WARNING: PDF is {int(file_age_seconds)}s old — may be from a previous location")

# ── MOVE AND RENAME ────────────────────────────────────────────────────
destination = os.path.join(maps_FOLDER, f"{normalized_address}_map.pdf")
log(f"Moving to {destination}")
try:
    shutil.move(newest_pdf, destination)
    log(f"SUCCESS: Moved to {destination}")
except Exception as e:
    log(f"ERROR during move: {type(e).__name__}: {e}")
    log(traceback.format_exc())
    raise SystemExit(1)
