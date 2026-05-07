import sys
import json
import openpyxl
from datetime import datetime
 
# ── INPUTS ────────────────────────────────────────────────────────────
# Power Automate passes two arguments:
#   sys.argv[1] = address (e.g. "11551 SHANNON DR")
#   sys.argv[2] = DocumentData JSON string
 
EXCEL_PATH = r"C:\Automation\portal_Compliance_Model.xlsx"
SHEET_NAME = "Raw_Document_Pull"
 
# ── VALIDATE ARGUMENTS ────────────────────────────────────────────────
try:
    with open(r'C:\Automation\temp_address.txt', 'r', encoding='utf-8-sig') as f:
        address = f.read().strip()
except FileNotFoundError:
    print('ERROR: temp_address.txt not found.')
    sys.exit(1)
 
TEMP_PATH = r"C:\Automation\temp_docs.json"
try:
    with open(TEMP_PATH, 'r', encoding='utf-8-sig') as f:
        raw_json = f.read().strip()
except FileNotFoundError:
    print(f"ERROR: Temp file not found at {TEMP_PATH}")
    sys.exit(1)
 
# ── PARSE JSON ────────────────────────────────────────────────────────
try:
    documents = json.loads(raw_json)
except json.JSONDecodeError as e:
    print(f"ERROR: Could not parse DocumentData JSON: {e}")
    sys.exit(1)
 
# ── CALCULATE EXP STATUS ──────────────────────────────────────────────
def get_exp_status(exp_date_str):
    if not exp_date_str or exp_date_str == "N/A":
        return "NO_EXP_REQUIRED"
    if exp_date_str == "01/01/1900":
        return "NO_EXP_REQUIRED"
    try:
        exp_date = datetime.strptime(exp_date_str, "%m/%d/%Y")
        today = datetime.today()
        days_until = (exp_date - today).days
        if days_until < 0:
            return "EXPIRED"
        elif days_until <= 30:
            return "EXPIRING_SOON"
        else:
            return "VALID"
    except ValueError:
        return "NO_EXP_REQUIRED"
 
# ── OPEN EXCEL AND WRITE ──────────────────────────────────────────────
try:
    wb = openpyxl.load_workbook(EXCEL_PATH)
except FileNotFoundError:
    print(f"ERROR: Excel file not found at {EXCEL_PATH}")
    sys.exit(1)
 
if SHEET_NAME not in wb.sheetnames:
    print(f"ERROR: Sheet '{SHEET_NAME}' not found in workbook.")
    sys.exit(1)
 
ws = wb[SHEET_NAME]
pull_date = datetime.today().strftime("%m/%d/%Y")
rows_written = 0
# ── DELETE EXISTING ROWS FOR THIS ADDRESS BEFORE REWRITING ────────────
rows_to_delete = []
 
for row_num in range(2, ws.max_row + 1):
    cell_value = ws.cell(row=row_num, column=1).value  # Column A = Address
    if cell_value and str(cell_value).strip().upper() == address.strip().upper():
        rows_to_delete.append(row_num)
 
for row_num in reversed(rows_to_delete):
    ws.delete_rows(row_num, 1)
 
for doc in documents:
    address_val      = address
    doc_area         = doc.get("table_type", "N/A")
    tech_name        = doc.get("employee_name", "N/A")
    doc_title        = doc.get("document", "N/A")
    exp_date         = doc.get("expiration", "N/A")
    location_specific = doc.get("location_specific", "N/A")
    inspection_date  = doc.get("inspection_date", "")
 
    # For qa_manual entries use inspection_date as a note in doc_title if present
    if doc_area == "qa_manual" and inspection_date:
        doc_title = doc_title  # keep as is — inspection_date stored separately
        tech_name = inspection_date  # store inspection date in tech_name field for qa_manual
 
    exp_status = get_exp_status(exp_date)
 
    ws.append([
        address_val,       # A - Address
        doc_area,          # B - doc_area
        tech_name,         # C - tech_name
        doc_title,         # D - doc_title
        exp_date,          # E - exp_date
        pull_date,         # F - pull_date
        exp_status,        # G - exp_status
        "PENDING",         # H - name_match (Python compliance layer will fill this)
        location_specific  # I - location_specific
    ])
    rows_written += 1
 
wb.save(EXCEL_PATH)
print(f"SUCCESS: {rows_written} rows written for {address}")