# Automation Flow

Step-by-step walkthrough of the Power Automate Desktop flows.

> Client and employer names are genericized throughout.

## Overview

There are two flows — one for Client A and one for Client B. They are near-identical in structure. The Client B flow was built by duplicating Client A and updating three references:

- Locations tab: `Client_A_Locations` → `Client_B_Locations`
- MultiDrop tab: `Client_A_MultiDrop` → `Client_B_MultiDrop`
- Dashboard click: Client A link → Client B link

Both run overnight. **Do not use the computer while a flow is running** — Power Automate controls Chrome, the keyboard, and the mouse.

Roughly 95 steps per flow. Everything from step 7 onward runs inside a For Each loop, once per location.

## Phase 1 — Setup (steps 1–6)

Runs once, before the loop. Loads everything into memory so Excel does not need to stay open.

| Step | Action |
|---|---|
| 1 | Launch Excel, open the compliance workbook, store as `ExcelInstance` |
| 2 | Activate the Locations tab for this client |
| 3 | Read column A into `AddressList` — this is what the loop iterates |
| 4 | Activate the MultiDrop tab |
| 5 | Read columns A and B into `MultiDropList` (A = address, B = dropdown number) |
| 6 | Close Excel — required, since Python cannot write to an open workbook |

> **Check the range in step 3 before every full run.** It is easy to leave a narrowed range behind after resuming a partial run, which silently skips every location above the start row.

## Phase 2 — Loop start and login (steps 7–13)

| Step | Action |
|---|---|
| 7 | For Each `CurrentAddress` in `AddressList` |
| 8 | Set `Address` from `CurrentAddress` |
| 9 | Write the address to `temp_address.txt` |
| 10 | Run `get_drop_number.py`, passing the client name so it checks the right MultiDrop tab. Must wait for completion or the next step reads an empty file |
| 11 | Read `temp_dropnumber.txt` into `DropNumber` (1, 2, or 3) |
| 12 | Launch Chrome at the portal login page, store as `Browser` |
| 13 | Display a message pausing for manual login and client selection |

## Phase 3 — Navigate to location (steps 14–26)

| Step | Action |
|---|---|
| 14 | Click the client account link on the dashboard |
| 15 | Wait for the Location search field |
| 16 | Click into the search field |
| 17 | Send keys — type the address |
| 18 | Wait for the suggestion dropdown |
| 19 | Click the first result |
| 20–26 | Multi-drop If/Else: `DropNumber = 2` → JavaScript clicks the second option; `= 3` → third; else proceed with the default |

> Steps in this phase that use recorded UI elements can capture the address text of whatever location was on screen when they were recorded. If a recorded click carries a literal address in its selector, it will only match that one location. Verify against a second address after any re-record.

## Phase 4 — Map detection (steps 27–38)

| Step | Action |
|---|---|
| 27 | Wait for the location homepage |
| 28 | Wait 15s for the Diagrams section to render |
| 29 | JavaScript: check for a floor plan PDF, set `MapExists` to YES/NO. Falls back to detecting "No available floor plans" text |
| 30 | Wait 15s |
| 31 | If `MapExists` = YES |
| 32 | JavaScript: format the address for use as a filename |
| 33 | JavaScript: click Print All, triggering the PDF download |
| 34 | Wait 5s for the download to start |
| 35 | Write the address to `temp_mapaddress.txt` so `move_map.py` knows the filename |
| 36 | Wait 20s for the download to finish |
| 37 | Run `move_map.py` — moves and renames the PDF into `Maps\` |
| 38 | End If |

## Phase 5 — Navigate to Licenses tab (steps 39–45)

| Step | Action |
|---|---|
| 39 | JavaScript: click the Digital Logbook link — more stable than a recorded UI element |
| 40 | Wait for the Licenses, Certifications and Insurance tab |
| 41 | Click into that tab |
| 42 | Wait for the Branch Contacts table |
| 43 | Wait 15s for table data to render |
| 44 | JavaScript: reset `window._auditResults` to empty and `window._auditComplete` to false — prevents the previous location's data carrying over |
| 45 | JavaScript: define scraping helpers — `cleanText()` and `getTableByHeading()` |

`getTableByHeading()` is the important one. It finds tables by heading text rather than DOM index, which is what makes the scraper survive portal UI changes.

## Phase 6 — Scrape four tables (steps 46–87)

| Table | Steps | Contents |
|---|---|---|
| Branch Contacts and Assigned Technicians | 46–55 | Currently assigned techs and their documents |
| Other Servicing Technicians | 56–65 | Techs who serviced in the past 12 months but are not assigned |
| Insurance and Certificates | 66–74 | COI and branch license — no employee name column |
| QA Inspection Reports | 75–87 | On the Reports and Logs tab: annual assessments, quarterly trends, pesticide logs |

Each table uses the same pagination pattern:

```
Set KeepPaging = true
Loop up to 20 iterations
    If KeepPaging = false → exit loop
    Wait 3 seconds
    Run JavaScript to read the next page
    Update KeepPaging
End loop
```

Columns captured: employee name, role, document name, document type, expiration date. For QA reports: inspection date and document name.

> Branch staff upload documents to whichever section is convenient. Python treats all four tables as one unified document pool — the source section is recorded as a hint, but never restricts what type of document is looked for where.

## Phase 7 — Save and write (steps 88–93)

| Step | Action |
|---|---|
| 88 | JavaScript: `JSON.stringify(window._auditResults)` → `DocumentData` |
| 89 | Write `DocumentData` to `temp_docs.json` |
| 90 | Write `CurrentAddress` to `temp_address.txt` |
| 91 | Run `write_to_excel.py` — appends one row per document to `Raw_Document_Pull`. Must wait for completion |
| 92 | Close the browser — a fresh one opens at step 12 for the next location |
| 93 | End For Each |

## Phase 8 — Post-loop (steps 94–95)

Runs once, after every location is done. This is the Python pipeline.

| Step | Action |
|---|---|
| 94a | `parse_service_detail.py` — rebuilds `WO_Counts_Detail` from raw exports, auto-deduping by work order number |
| 94b | `compare_equipment.py` — rebuilds `Equipment_Compliance` |
| 95 | `evaluate_compliance.py` — reads everything, writes `Compliance_Eval` |

> These three can be run standalone from the command line without re-scraping. That makes a mid-month equipment check cheap: refresh the raw exports, run 94a → 94b → 95, done.

## Known constraints

**Power Fx is enabled and cannot be disabled** without rebuilding both flows from scratch. Consequences:

- Numbers need an `=` prefix: `=1`, `=20`
- String comparisons use single quotes
- Boolean logic needs Set-variable workarounds with `If()`
- **Variables must always be inserted with the `{x}` picker.** Typing a variable name manually writes it as literal text instead of its value

**Recorded UI elements are fragile.** Power Automate identifies elements by their attributes, and portal updates change them. Re-record the element when a click breaks. JavaScript steps that target by heading text are more resilient and are preferred for anything load-bearing.

**Multi-drop locations.** Roughly 17 locations return multiple results when searched. These are flagged `MANUAL_REVIEW` in `Compliance_Eval` rather than evaluated automatically. A dedicated flow for them is deferred.

**Python must be invoked as `py`, not `python`,** using the full path to the launcher — Power Automate does not inherit PATH.
