# Excel Structure

What every tab and column in both workbooks means.

> Client and employer names are genericized throughout.

## Overview

Two workbooks power the pipeline. They are linked by address — Python normalizes addresses from both before joining.

| File | Purpose |
|---|---|
| `Compliance_Model.xlsx` | Compliance brain — location lists, scraped documents, results |
| `Service_Ticket_Report.xlsx` | Work-order master — contract scope, work orders, equipment verdicts |

Both live on a local path, never in a cloud-synced folder.

---

# Compliance_Model.xlsx

Tabs fall into three categories: **reference** (static lookups), **staging** (refreshed each run), and **output** (written by Python).

## Reference tabs

### `Client_A_Locations` / `Client_B_Locations`

~131 and ~188 addresses respectively. Same structure.

| Column | Content |
|---|---|
| A | Address |
| C | Company / location type |
| D | Branch number |
| E | State |
| F | Franchise flag (YES/NO) |

A branch number starting with `9` is also treated as a franchise, regardless of the flag.

### `Client_A_MultiDrop` / `Client_B_MultiDrop`

Addresses that return multiple results when searched in the portal.

| Column | Content |
|---|---|
| A | Address |
| B | Dropdown number to select (2 or 3) |

A value of 2 or 3 means the correct result is not the first one, so the scrape cannot be trusted. Those locations are marked `MANUAL_REVIEW`. Drop 1 or blank is an ordinary location.

### `State_Exceptions`

States where individual technician licenses are not required. Python assigns `EXEMPT` to the tech license check for locations in those states.

| Column | Content |
|---|---|
| A | Two-letter state code |
| B | Required? YES / NO |

### `Doc_Rules`

Document classification rules. Informational — the live rules are in `evaluate_compliance.py`.

## Staging tabs

### `Raw_Document_Pull`

Every document scraped by Power Automate. One row per document per location per pull date.

| Col | Field | Content |
|---|---|---|
| A | address | Service location address |
| B | doc_area | Which table section the document came from |
| C | tech_name | Employee name — or, for QA rows, the inspection date |
| D | doc_title | Document name exactly as it appears in the portal |
| E | exp_date | Expiration date if present |
| F | pull_date | Date this row was scraped |
| G | exp_status | VALID / EXPIRED / EXPIRING_SOON / NO_EXP_REQUIRED |
| H | name_match | `PENDING` until `evaluate_compliance.py` runs |
| I | location_specific | Whether the document is specific to this location |

Two behaviors worth knowing:

- **Inspection dates are stored in `tech_name` for QA rows.** This is deliberate — QA rows have no employee, and `evaluate_compliance.py` reads that column differently based on `doc_area`.
- **History is preserved.** Re-running an address on the same day replaces that day's rows; previous days stay. `evaluate_compliance.py` filters to the most recent pull per address, so old rows never affect a verdict.

### `Map_Manual_Data`

Manually verified map device counts. **Python only ever writes columns A and B.** Everything else is yours and is never overwritten.

| Column | Updated by | Content |
|---|---|---|
| address | Python | Auto-populated for new locations |
| map_found | Python | YES / NO |
| map_erb | Manual | Verified exterior bait station count |
| map_irt | Manual | Verified interior rodent trap count |
| map_ifl | Manual | Verified fly light count |
| map_dated | Manual | Is the map dated? YES / NO |
| map_signed | Manual | Is the map signed? YES / NO |
| map_status | Manual | Overall status |
| last_verified | Manual | Date last checked |
| map_flags | Manual | Notes and issues |

## Output tab

### `Compliance_Eval`

One row per location per month. 25 columns.

| Col | Field | Content |
|---|---|---|
| A | audit_month | e.g. `August 2026` — set by `AUDIT_MONTH` in the script |
| B | address | Display format |
| C | branch | Branch number |
| D | state | Two-letter code |
| E | company | Client and location type |
| F | tech_name | Earliest non-zero work-order tech for this address and month |
| G | PC_license_branch | PASS / EXPIRED / MISSING / MANUAL_REVIEW |
| H | PC_license_tech | PASS / EXPIRED / MISSING / EXEMPT / MANUAL_REVIEW |
| I | client_cert | PASS / MISSING / MANUAL_REVIEW |
| J | IPM_cert | PASS / EXPIRED / MISSING / WRONG_YEAR / MANUAL_REVIEW |
| K | COI | PASS / EXPIRED / MISSING / WRONG_YEAR / MANUAL_REVIEW |
| L | annual_report | PASS / MISSING / MANUAL_REVIEW |
| M | quarterly_trend | PASS / MISSING / EXEMPT / MANUAL_REVIEW |
| N | pesticide_log | PASS / MISSING / EXEMPT / MANUAL_REVIEW |
| O | equipment_match | From `Equipment_Compliance` — PASS / UNDER / OVER / CANCELLED / CHECK / NO_DATA |
| P | map_found | YES / NO |
| Q | map_erb | Verified count from the map |
| R | map_irt | Verified count from the map |
| S | map_ifl | Verified count from the map |
| T | map_dated | YES / NO |
| U | map_signed | YES / NO |
| V | map_vs_sow | MATCH / MISMATCH / NO_MAP / NO_SOW / NO_MAP_COUNTS / MANUAL_REVIEW |
| W | map_mismatch_detail | e.g. `ERB: Map=34 SOW=32` — blank when MATCH |
| X | flags | All issues combined into one column for filtering |
| Y | last_checked | Date the script wrote this row |

**Write behavior.** Rows are keyed on `(address, audit_month)`. A matching row is overwritten; a new one is appended. Duplication can be allowed deliberately as a reference when the portal is suspected of glitching, so two pulls can be compared side by side.

> A compliance percentage is calculated during the run and printed to the console, but is **not** written as a column. If a percentage is needed in reporting, it can be computed downstream from the eight check columns — count `PASS` divided by count of non-`EXEMPT`.

---

# Service_Ticket_Report.xlsx

Four tabs.

### `SOW_Lookup` — built by `read_sow_v2.py`

One row per location: contracted scope. Wiped and replaced on each run of that script.

| Col | Field |
|---|---|
| A | Address |
| B–E | ERB: device count, frequency, visits per month, expected monthly scans |
| F–I | IRT: same four fields |
| J–M | IFL: same four fields |
| N | Flags — includes the `CANCELLED` marker |

Frequency translates to visits per month as: monthly = 1, semi-monthly / every other week = 2, weekly = 4. Expected scans = device count × visits per month.

### `WO_Counts_Detail` — built by `parse_service_detail.py`

One row per unique work order. Wiped and replaced each run.

| Col | Field |
|---|---|
| A | WORK ORDER |
| B | ADDRESS (cleaned) |
| C | RAW_ADDRESS (as it appeared in the export) |
| D | TECHNICIAN |
| E | BRANCH |
| F | COMPLETED |
| G | MONTH — e.g. `Aug 2026` |
| H | CLIENT |
| I | ACTIVITY |
| J–L | ERB / IRT / IFL scan counts |
| M | REVIEW_FLAGS |

`REVIEW_FLAGS` carries anything the parser could not resolve confidently — an ambiguous zone, an unrecognized device type, or a new address with no contract row yet.

### `Equipment_Compliance` — built by `compare_equipment.py`

One row per location per month. Wiped and replaced each run. 24 columns.

Identity: `ADDRESS`, `CLIENT`, `MONTH`, `MATCH_TYPE` (exact / fuzzy / none).

Then six columns for each of the three device types:

| Suffix | Meaning |
|---|---|
| `_DEVICES` | Devices on site, from the contract |
| `_FREQUENCY` | How often each should be serviced |
| `_EXPECTED_SCANS` | Devices × visits per month |
| `_TOTAL_SCANS` | Actual scans summed across the month's work orders |
| `_EFFECTIVE_VISITS` | Total scans ÷ devices — average visits per device |
| `_VERDICT` | PASS / UNDER / OVER / CANCELLED / EXTRA DEVICE / NO SOW / N/A |

Then `OVERALL_VERDICT` and `FLAGS`.

`_EFFECTIVE_VISITS` is the column a branch manager actually reads — it converts an abstract scan count into "your tech visited each device about 3 times this month when the contract says 2."

**Verdict thresholds:** within 95–110% of expected is `PASS`; below 95% is `UNDER (n short)`; above 110% is `OVER (~Nx, expected Mx)`.

Locations in `SOW_Lookup` with no work orders at all get a row marked `NO WORK ORDERS — verify (cancelled?)`, so a site that quietly stopped being serviced cannot disappear from the report.

### `Location_Master`

Address → branch and company lookup. Static reference.

---

## Critical rules

- Close Excel completely before running any Python script — `openpyxl` cannot write to an open file
- Set `AUDIT_MONTH` in `evaluate_compliance.py` before every run
- Map count columns in `Map_Manual_Data` are manual — Python never overwrites them
- `SOW_Lookup`, `WO_Counts_Detail`, and `Equipment_Compliance` are all wiped and rebuilt by their scripts. Do not hand-edit them
- Never move either workbook into a cloud-synced folder
