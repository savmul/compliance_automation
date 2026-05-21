# Excel Structure

## Overview

Two Excel files power this pipeline.
They are linked by ADDRESS — Python 
normalizes addresses from both files 
before joining them.

| File | Purpose |
|---|---|
| Insite_Compliance_Model.xlsx | Compliance brain — documents, results, location lists |
| Service_Ticket_Report.xlsx | Work order master — raw service data |

All files stored at C:\Automation\
Never move to a cloud-synced folder.

---

## Insite_Compliance_Model.xlsx

### Tab Types

There are three categories of tabs:

| Type | Description |
|---|---|
| Reference tabs | Static lookup data — do not edit manually |
| Staging tabs | Replaced or updated each month |
| Output tab | Written by Python — compliance results |

---

### Reference Tabs

**Client_A_Locations**
~131 addresses for Client A.

| Column | Content |
|---|---|
| A | Address |
| C | Company |
| D | Branch number |
| E | State |
| F | Franchise flag |

**Client_B_Locations**
~188 addresses for Client B.
Same column structure as Client_A_Locations.

**Client_A_MultiDrop**
Addresses that return multiple dropdown 
options when searched in the portal.

| Column | Content |
|---|---|
| A | Address |
| B | Dropdown number to select (2 or 3) |

**Client_B_MultiDrop**
Client B equivalent of Client_A_MultiDrop.

**State_Exceptions**
States where individual technician 
licenses are not required.
Python reads this tab to assign 
EXEMPT status to affected locations.

---

### Staging Tabs

**Raw_Document_Pull**
All documents scraped by Power Automate 
land here. One row per document 
per location.

Cleared before each full flow run to 
prevent old data mixing with new.

Key columns:

| Column | Content |
|---|---|
| address | Service location address |
| pull_date | Date this row was scraped |
| tech_name | Employee name (or inspection date for qa_manual rows) |
| doc_area | Which table section the doc came from |
| doc_title | Full document name as it appears in portal |
| doc_type | Certificate / License / Miscellaneous |
| expiry_date | Expiration date if present |
| exp_status | VALID / EXPIRED / EXPIRING_SOON / NO_EXP_REQUIRED |
| name_match | PENDING until evaluate_compliance.py runs |

> Note: For qa_manual rows the inspection 
> date is stored in the tech_name column 
> by design. evaluate_compliance.py 
> handles these rows differently.

**Last_WO_Tech**
Most recent work order and technician 
name per address. ~319 rows.
Replaced each month.
Python reads this to know whose 
documents to check for tech-level 
compliance checks.

| Column | Content |
|---|---|
| address | Service location address |
| last_wo_date | Date of most recent work order |
| tech_name | Technician who performed the work |

**Map_Manual_Data**
Manually verified map device counts 
per location.
The map_found column is auto-updated 
by Python each run.
Count columns are never overwritten 
by Python — only updated manually.

| Column | Updated By |
|---|---|
| address | Python — auto-populated |
| map_found | Python — YES / NO / MISSING |
| map_erb | Manual — verified ERB count |
| map_irt | Manual — verified IRT count |
| map_ifl | Manual — verified IFL count |
| map_dated | Manual — is the map dated? |
| map_signed | Manual — is the map signed? |
| map_status | Manual — overall status |
| last_verified | Manual — date last checked |
| map_flags | Manual — notes and issues |

---

### Output Tab

**Compliance_Eval**
One row per location per month.
Written by evaluate_compliance.py.
Currently wipes and rewrites on 
every run — will switch to append 
mode before go-live.

Key columns:

| Column | Content |
|---|---|
| audit_month | Month being evaluated |
| address | Service location address |
| branch | Branch number |
| state | State code |
| company | Client and location type |
| tech_name | Technician from Last_WO_Tech |
| PC_license_branch | PASS / MISSING / EXPIRED / MANUAL_REVIEW |
| PC_license_tech | PASS / MISSING / EXPIRED / EXEMPT / MANUAL_REVIEW |
| COI | PASS / MISSING / EXPIRED / WRONG_YEAR / MANUAL_REVIEW |
| pepsi_cert | PASS / MISSING |
| IPM_cert | PASS / MISSING / EXPIRED / WRONG_YEAR / MANUAL_REVIEW |
| annual_report | PASS / MISSING / MANUAL_REVIEW |
| quarterly_trend | PASS / MISSING / EXEMPT / MANUAL_REVIEW |
| pesticide_log | PASS / MISSING / EXEMPT / MANUAL_REVIEW |
| equipment_match | MATCH / MISMATCH / NO_DATA |
| map_found | YES / NO / MISSING |
| map_erb | Verified ERB count from map |
| map_irt | Verified IRT count from map |
| map_ifl | Verified IFL count from map |
| map_signed | YES / NO |
| map_dated | YES / NO |
| checks_applicable | Number of checks that apply |
| checks_passed | Number of checks that passed |
| compliance_pct | checks_passed ÷ checks_applicable |
| flags | Pipe-separated list of failed checks |
| last_checked | Timestamp of last evaluation run |

---

## Service_Ticket_Report.xlsx

### Physical Tabs

**page**
Raw work order data. All cleaned monthly 
rows paste here.

> CRITICAL RULES FOR THIS TAB:
> - Never rename this tab — Power Query 
>   is hardcoded to this name
> - Never paste the header row
> - Always paste values only — 
>   never with formatting
> - Never edit Power Query unless 
>   intentionally rebuilding

Column order — 9 columns in exact order:

| Column | Name | Notes |
|---|---|---|
| A | WORK ORDER | Unique ticket number |
| B | COMPANY | Client and location type |
| C | ADDRESS | Service location address |
| D | TECHNICIAN | Tech who performed service |
| E | SERVICE TYPE | e.g. PC Standard - Monthly |
| F | DATETIME | Timestamp of service |
| G | ZONE NAME | Zone where equipment is located |
| H | TRAP TYPE | Standardized to 3 values only |
| I | STATION FINDING | Observation Recorded / Pesticide Usage / Pest Finding / blank |

**Location_Master**
Branch, address, and company lookup.
Used by Power Query to merge branch 
numbers into Location_Month_Summary.
Do not edit manually.

---

### Power Query Tabs

Run automatically on Refresh All.
Do not edit unless intentionally rebuilding.

| Query | What It Does |
|---|---|
| tbl_Raw | Reads page tab, adds calculated columns including ScanCount and equipment flags |
| WO_Summary | Groups by work order, sums ERB/IRT/IFL counts and zone observations |
| Location_Month_Summary | Groups to location level, merges with Location_Master for branch data |
| tbl_Location_Master | Reads Location_Master tab as clean lookup |

> The tab Python reads must be named 
> exactly Location_Month_Summary — 
> no variation. Python will error 
> if the name does not match.

---

## Trap Type Standardization

Column H in the page tab must contain 
exactly one of these three values.
Standardize using Find and Replace 
before pasting each month.

| Standard Value | What It Represents |
|---|---|
| Ext Rodent Bait Station | Exterior rodent bait stations |
| Tin Cat Trap | Interior rodent traps |
| Fly Light | Insect and fly light traps |

---

## Equipment Count Logic

| Finding Value | Counts As |
|---|---|
| Observation Recorded | 1 scanned device |
| Pesticide Usage | 0 — does not count as scan |
| Pest Finding | 0 — does not count as scan |
| blank | 0 — equipment not scanned |

If a device has both Observation Recorded 
and Pesticide Usage in the same work order, 
Power Query deduplicates automatically — 
it counts as 1 scan, not 2.
