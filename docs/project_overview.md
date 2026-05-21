# Project Overview

## Summary

An end-to-end compliance automation pipeline built 
to replace a fully manual monthly audit process 
across ~319 service locations for two corporate 
accounts.

A Power Automate Desktop robot scrapes a customer 
portal nightly, Python evaluates compliance against 
8 checks per location, and results are written to 
Excel for review and reporting.

---

## The Problem

Each service location must maintain active 
compliance documents at all times:

- Branch pest control business license
- Certificate of Insurance (COI)
- Technician pesticide license
- Technician client certification
- IPM / cGMP certification
- Annual facility assessment
- Quarterly pest trend report
- Quarterly pesticide usage log

With ~319 locations across two accounts, manually 
checking 8 documents per location every month is 
not sustainable. Documents expire, names are 
inconsistent, some belong to specific technicians, 
and some require current-year versions only.

This pipeline automates collection and evaluation 
of all compliance data monthly.

---

## Architecture

Portal → Power Automate → temp files →
Python → Excel → Power BI (in progress)

**Data flows in two streams that meet in Python:**

Stream 1 — Work Orders
Portal export → manual cleaning →
Excel (page tab) → Power Query →
Location_Month_Summary

Stream 2 — Compliance Documents
Portal → Power Automate scrapes each location →
write_to_excel.py → Raw_Document_Pull

Both streams join in evaluate_compliance.py, 
which runs 8 checks per location and writes 
results to Compliance_Eval.

---

## Scale

| Item | Value |
|---|---|
| Client A locations | ~131 |
| Client B locations | ~188 |
| Total locations | ~319 |
| Compliance checks per location | 8 |
| Automation flows | 2 (one per client) |
| Run frequency | Monthly |

---

## Tools & Technologies

| Tool | Purpose |
|---|---|
| Power Automate Desktop | Browser automation and portal scraping |
| Python 3 | Compliance evaluation and data processing |
| JavaScript | Injected via PAD to scrape portal page tables |
| openpyxl | Python library for reading and writing Excel |
| Excel | Central data store |
| Power Query | Auto-summarizes work order data on refresh |
| Power BI | Planned reporting layer — pending licensing |

---

## Key Engineering Decisions

**Temp files as communication layer**
Power Automate cannot reliably pass variables 
to Python as command line arguments — special 
characters in addresses cause crashes. All 
data passes through temp files instead 
(temp_address.txt, temp_docs.json, 
temp_dropnumber.txt).

**Heading-based table targeting**
JavaScript scraping uses getTableByHeading() 
to find tables by their heading text rather 
than hardcoded index numbers. This makes the 
scraper resilient to portal UI updates.

**Hash-based map change detection**
Floor plan PDFs are manually verified once and 
stored in Map_Manual_Data. map_hash.py 
fingerprints each PDF monthly — only changed 
maps require re-review.

**Typo normalization**
Document names uploaded by branch staff are 
inconsistent. A TYPO_MAP dictionary 
auto-corrects known misspellings before 
keyword matching runs.

**OCR abandoned**
Tesseract OCR was tested for reading device 
counts from map PDFs. Results were unreliable 
due to inconsistent map quality across 
locations. Replaced with manual verification 
plus hash-based change detection.

---

## Compliance Check Results

| Value | Meaning |
|---|---|
| PASS | Document found, valid, correct year, correct tech |
| MISSING | No matching document found |
| EXPIRED | Document found but expiration date has passed |
| WRONG_YEAR | Document is from prior year |
| MANUAL_REVIEW | Unclear date/name or multi-drop location |
| EXEMPT | Check does not apply to this location |
| NO_DATA | Work order data not found for this period |

---

## Current Status

| Component | Status |
|---|---|
| Client A automation flow | Complete |
| Client B automation flow | Complete |
| write_to_excel.py | Complete |
| evaluate_compliance.py | Complete |
| Full monthly runs | Complete |
| Last_WO_Tech auto-update script | In progress |
| Power BI dashboard | Pending licensing |

---

## Author

Savanna Mullins
Built independently as a self-initiated project 
outside of a formal data role.
