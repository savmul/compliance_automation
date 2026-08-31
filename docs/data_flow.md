# Data Flow

How data moves from the portal to the final report.

> Client and employer names are genericized throughout.

## Overview

Three independent streams feed the final verdict. All three join on a normalized address key, so as long as a location has data in each stream, the verdict ties together.

| Stream | Source | Lands in | Read by |
|---|---|---|---|
| Work orders & equipment | Raw "Service Detail by Workorder" exports | `WO_Counts_Detail`, `Equipment_Compliance` | `evaluate_compliance.py` |
| Documents | Power Automate scrape of the Licenses and Reports tabs | `Raw_Document_Pull` | `evaluate_compliance.py` |
| Contract scope & maps | Master Grid via `read_sow_v2.py`, plus manual map counts | `SOW_Lookup`, `Map_Manual_Data` | `compare_equipment.py`, `evaluate_compliance.py` |

Only the document stream touches the portal's user interface. The other two read exported files, so a portal front-end change cannot break them.

## Address normalization — the join key

Every script normalizes addresses in memory before joining. The stored values are never modified.

The same logic appears in `parse_service_detail.py`, `compare_equipment.py`, `evaluate_compliance.py`, and `write_to_excel.py`. **If you change one, change all of them** — they must agree on what counts as the same location.

The rules:

1. Keep only the portion before the first comma (drops city, state, ZIP)
2. Uppercase, and remove periods
3. Collapse repeated whitespace
4. Drop any leading text before the first token that starts with a digit — this strips location-type prefixes like `CLIENT A PLANT 1234 MAIN ST` down to `1234 MAIN ST`
5. Standardize street suffixes: `ROAD`→`RD`, `STREET`→`ST`, `AVENUE`→`AVE`, `DRIVE`→`DR`, `BOULEVARD`→`BLVD`, and similar

`compare_equipment.py` additionally falls back to fuzzy matching at a 0.90 similarity cutoff when no exact key match exists, and marks those rows `FUZZY MATCH — verify` so they can be reviewed.

## Stream 1 — Work orders and equipment

**Source.** The "Service Detail by Workorder" report, exported from the portal for both clients into `Raw Data\`. Filenames must begin with the client prefix; anything after it is free, so date ranges can be written however is convenient.

**`parse_service_detail.py`** reads every matching file and, for each work order:

- Locates header fields by label text rather than fixed column position, because the report shifts label columns between exports
- Filters to standard recurring service. First-service work orders are kept only for addresses that have never had a standard service, and are flagged `NEW ADDRESS — needs SOW setup`
- Categorizes each scanned device into one of three types: exterior rodent bait stations (ERB), interior rodent traps (IRT), or insect/fly lights (IFL)
- Uses zone description as a tiebreaker for device types whose names are ambiguous, flagging anything still unclear
- Deduplicates devices by barcode within a work order, so a station scanned twice counts once
- Deduplicates work orders by ticket number across all files, so overlapping date ranges are safe

Output: `WO_Counts_Detail`, one row per unique work order.

**`compare_equipment.py`** then groups those work orders by location and month, joins to `SOW_Lookup`, and compares actual scans to `devices × visits-per-month`.

Output: `Equipment_Compliance`, one row per location per month, with a verdict per device type and an overall verdict.

## Stream 2 — Compliance documents

**Source.** The portal's Licenses, Certifications and Insurance tab plus the Reports and Logs tab, scraped for every location.

Four table sections are collected:

| Section | Documents typically found |
|---|---|
| Branch Contacts and Assigned Technicians | Technician licenses and certifications |
| Other Servicing Technicians | Documents for technicians who serviced in the past 12 months |
| Insurance and Certificates | COI, branch license |
| QA Inspection Reports | Annual assessments, quarterly trend reports, pesticide usage logs |

Branch staff upload documents to whichever section is convenient — there is no enforcement. Python therefore treats all four sections as **one unified document pool**. The source section is recorded and used as a disambiguation hint, but it never restricts what type of document can be found where.

**Temp file communication layer.** Power Automate cannot reliably pass variables to Python as command-line arguments, so every hand-off is a file.

| Temp file | Written by | Read by |
|---|---|---|
| `temp_address.txt` | Power Automate | `write_to_excel.py`, `get_drop_number.py` |
| `temp_docs.json` | Power Automate | `write_to_excel.py` |
| `temp_dropnumber.txt` | `get_drop_number.py` | Power Automate |
| `temp_mapaddress.txt` | Power Automate | `move_map.py` |

`write_to_excel.py` reads these with an encoding-tolerant reader — UTF-8-BOM, then UTF-16, then Latin-1 — because Power Automate is inconsistent about which it writes.

Output: `Raw_Document_Pull`, one row per document per location per pull date.

**History is preserved.** Re-running the same address on the same day replaces that day's rows. Previous days' pulls are left in place, so day-over-day portal discrepancies stay visible. `evaluate_compliance.py` filters to the most recent pull per address, so the extra history never affects a verdict.

## Stream 3 — Contract scope and maps

**`read_sow_v2.py`** reads the Master Grid and writes `SOW_Lookup`: device count, service frequency, visits per month, and expected monthly scans per device type, plus a cancelled marker. Run manually, only when contract scope changes.

**Map counts** are verified by hand once per location and entered in `Map_Manual_Data`. `map_hash.py` fingerprints each map PDF monthly, so only changed maps need re-review. Python never overwrites the count columns.

## Merge point — `evaluate_compliance.py`

```
Raw_Document_Pull ──┐
Equipment_Compliance ┤
SOW_Lookup ──────────┼──► evaluate_compliance.py ──► Compliance_Eval
WO_Counts_Detail ────┤
Map_Manual_Data ─────┤
Location tabs ───────┘
```

Per location it:

1. Identifies the responsible technician — the earliest non-zero work order of the audit month, with `LAST, FIRST` flipped to `FIRST LAST` for name matching
2. Runs the 8 document checks against that technician's documents
3. Pulls the pre-computed equipment verdict from `Equipment_Compliance` (single source of truth — no recomputation)
4. Compares verified map counts to contracted counts per device type
5. Writes one row to `Compliance_Eval`

**Write behavior.** Rows are keyed on `(address, audit_month)`. A matching row is overwritten; a new one is appended. Duplication can be allowed intentionally as a reference when the portal is suspected of glitching, so day-over-day differences can be compared.

## Supporting flows

**Map PDFs.** Portal → flow detects a Diagrams section → JavaScript clicks Print All → PDF downloads → `move_map.py` renames it by normalized address and files it in `Maps\` → `map_hash.py` detects changes month over month.

**Drop number lookup.** Flow starts a location → `get_drop_number.py` checks the MultiDrop tab → writes 1, 2, or 3 to `temp_dropnumber.txt` → flow reads it and selects the correct search result.

## File locations

All working files live on a local path, never in a cloud-synced folder — sync conflicts caused file-locking errors during overnight runs.

| File | Purpose |
|---|---|
| `Compliance_Model.xlsx` | Compliance workbook — locations, documents, results |
| `Service_Ticket_Report.xlsx` | Work-order master — scope, work orders, equipment verdicts |
| `Raw Data\` | Folder for raw portal exports |
| `Maps\` | Map PDFs, named by normalized address |

## Critical rules

- Close Excel completely before running any Python script — `openpyxl` cannot write to an open file
- Set `AUDIT_MONTH` in `evaluate_compliance.py` before every run
- Raw data filenames must begin with the client prefix or they will not be found
- Run `read_sow_v2.py` only when contract scope changes
- Map count columns in `Map_Manual_Data` are manual — the script never overwrites them
- Never move working files into a cloud-synced folder
- Address normalization logic must stay identical across all scripts
