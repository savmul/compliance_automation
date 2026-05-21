# Data Flow

## Overview

Data moves through this pipeline in two separate 
streams that merge in Python before results are 
written to Excel.

STREAM 1 — WORK ORDERS
Portal export → manual cleaning → Excel →
Power Query → Location_Month_Summary
STREAM 2 — COMPLIANCE DOCUMENTS
Portal → Power Automate → temp files →
Python → Raw_Document_Pull
MERGE POINT
evaluate_compliance.py joins both streams
on ADDRESS → runs 8 checks → writes
Compliance_Eval

---

## Stream 1 — Work Orders

**Source**
Monthly service ticket report exported 
manually from the portal.

**Manual cleaning steps (required before paste)**
- Delete unnecessary columns from raw export
- Standardize Trap Type to exactly 3 values:
  - Ext Rodent Bait Station
  - Tin Cat Trap  
  - Fly Light
- Remove non-standard service rows
- Delete header row before pasting

**Destination**
Cleaned rows pasted into the page tab of 
the work order Excel file.
Values only — never paste with formatting.
Never paste the header row.

**Power Query transforms (automatic on Refresh All)**

| Query | What It Does |
|---|---|
| tbl_Raw | Reads page tab, adds calculated columns |
| WO_Summary | Groups by work order, sums equipment counts |
| Location_Month_Summary | Groups to location level, merges branch data |
| tbl_Location_Master | Clean address and branch lookup |

**What Python reads from Stream 1**
Location_Month_Summary tab — provides:
- SOW equipment counts (ERB, IRT, IFL)
- Actual scanned equipment counts
- Branch number and company per location

---

## Stream 2 — Compliance Documents

**Source**
Portal — Licenses, Certifications and 
Insurance tab plus Report and Logs tab, 
scraped for every location.

**Collection**
Power Automate Desktop navigates to each 
location and injects JavaScript to scrape 
four table sections:

| Section | Documents Found |
|---|---|
| Branch Contacts and Assigned Technicians | Tech licenses, certifications |
| Other Servicing Technicians | Documents for past techs |
| Insurance and Certificates | COI, branch license |
| QA Inspection Reports | Annual reports, quarterly logs |

**Temp file communication layer**
Power Automate writes scraped data to 
temp files. Python reads from those files.
Direct variable passing is unreliable 
due to special characters in addresses.

| Temp File | Written By | Read By |
|---|---|---|
| temp_address.txt | Power Automate | write_to_excel.py, get_drop_number.py |
| temp_docs.json | Power Automate | write_to_excel.py |
| temp_dropnumber.txt | get_drop_number.py | Power Automate |
| temp_mapaddress.txt | Power Automate | move_map.py |

**What Python reads from Stream 2**
Raw_Document_Pull tab — provides:
- Document names per location
- Expiration dates
- Technician names from document titles
- Document source table (for reference)

---

## Merge Point — evaluate_compliance.py

Python joins both streams on ADDRESS and 
runs 8 compliance checks per location.

Raw_Document_Pull  ──┐
├──► evaluate_compliance.py ──► Compliance_Eval
Location_Month_     ──┘
Summary
Last_WO_Tech       ──┘
Pepsi_Locations    ──┘
Frito_Locations    ──┘
Map_Manual_Data    ──┘

Address normalization is applied before 
joining — strips punctuation, uppercases, 
collapses spaces — so minor formatting 
differences do not cause missed matches.

---

## Output — Compliance_Eval

One row per location per month.

| Column Group | What It Contains |
|---|---|
| Identity | Address, branch, state, company, audit month |
| 8 compliance checks | PASS / MISSING / EXPIRED / WRONG_YEAR / MANUAL_REVIEW / EXEMPT |
| Equipment match | SOW vs actual counts for ERB, IRT, IFL |
| Map status | map_found, map_erb, map_irt, map_ifl |
| Summary | compliance_pct, flags, checks_passed, last_checked |

---

## Supporting Data Flows

**Map PDFs**
Portal → Power Automate detects map →
JavaScript clicks Print All →
PDF downloads → move_map.py renames
and moves to Maps\ folder →
map_hash.py detects changes monthly
**Drop number lookup**
Flow starts new location →
get_drop_number.py checks MultiDrop tab →
writes result to temp_dropnumber.txt →
flow reads file → selects correct
portal dropdown option

**Technician lookup**
Last_WO_Tech tab → most recent tech
per address → Python cross-references
tech name against document titles →
tech-level checks evaluated

---

## File Locations

All files stored at C:\Automation\ — 
not in any cloud-synced folder.
Cloud sync conflicts caused file locking 
errors during overnight runs.

| File | Purpose |
|---|---|
| Insite_Compliance_Model.xlsx | Main compliance workbook |
| 2026_Service_Ticket_Report.xlsx | Work order master file |
| evaluate_compliance.py | Monthly compliance evaluation |
| write_to_excel.py | Writes per-location doc data |
| get_drop_number.py | Dropdown lookup per location |
| move_map.py | Moves and renames map PDFs |
| map_hash.py | Detects changed maps monthly |
| Maps\ | Folder of all map PDFs by address |

---

## Critical Rules

- Never rename the page tab in the 
  work order file
- Never paste the header row into 
  the page tab
- Always paste values only — 
  never with formatting
- Always close Excel completely 
  before running Python
- Always update AUDIT_MONTH in 
  evaluate_compliance.py before 
  each monthly run
- Confirm Location_Month_Summary 
  tab name is exact before running
- Never move files from C:\Automation\ 
  to a cloud-synced folder
