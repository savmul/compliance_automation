# Project Overview

Plain-English explanation of what this system is and why it is built the way it is.

> Client and employer names are genericized throughout. "Client A" and "Client B" are two national food & beverage accounts. "The portal" is a proprietary service-management platform.

## Summary

An end-to-end compliance automation pipeline replacing a fully manual monthly audit across 319 service locations for two corporate accounts.

A Power Automate Desktop robot scrapes the customer portal location by location. Python then reconciles three independent data sources, evaluates 8 document checks plus equipment and map comparisons per location, and writes structured verdicts to Excel.

## The problem

Each location must maintain current compliance documentation at all times:

1. Branch pest control business license
2. Certificate of Insurance (COI)
3. Technician pesticide license
4. Technician client certification
5. IPM / cGMP certification
6. Annual facility assessment
7. Quarterly pest trend report
8. Quarterly pesticide usage log

At 319 locations that is 2,552 document verifications per month, done by hand.

The harder problem is that the documents are not clean data. Titles are free text entered by branch staff, so the same license appears a dozen different ways and is frequently misspelled. Staff upload documents to whichever section of the portal is convenient, so section is a hint rather than a rule. Recorded expiration dates are wrong roughly 15% of the time. Some documents belong to a specific technician and only count if that technician is the one currently servicing the site.

Separately, every location has a contracted device count and service frequency. Confirming that a technician actually serviced the right devices the right number of times means reconciling work-order exports against a contract grid — a step that was rarely done manually because it was too laborious.

## Architecture

Three data streams join on a normalized address key.

**Stream 1 — Work orders and equipment**
Raw "Service Detail by Workorder" exports → `parse_service_detail.py` → `WO_Counts_Detail`

**Stream 2 — Compliance documents**
Portal → Power Automate scrape → temp files → `write_to_excel.py` → `Raw_Document_Pull`

**Stream 3 — Contract scope and maps**
Master Grid → `read_sow_v2.py` → `SOW_Lookup`, plus manually verified counts in `Map_Manual_Data`

`compare_equipment.py` joins streams 1 and 3 to produce `Equipment_Compliance`. `evaluate_compliance.py` then joins everything and writes `Compliance_Eval`.

Only Stream 2 depends on the portal's user interface. That separation is deliberate: when the portal was rebuilt on a new front-end framework, streams 1 and 3 were unaffected and equipment reporting kept running while the scraper was repaired.

## Scale

| Item | Value |
|---|---|
| Client A locations | ~131 |
| Client B locations | ~188 |
| Total locations | 319 |
| Document checks per location | 8 |
| Device types reconciled | 3 |
| Automation flows | 2 (one per client) |
| Locations requiring manual review by design | ~17 |
| Run cadence | Monthly |

## Key engineering decisions

**Verdicts instead of booleans.** The original equipment check returned MATCH / MISMATCH and was wrong constantly on weekly-service locations, where 54 devices serviced weekly produce ~216 scans a month and any exact-match test fails. The current version compares actual scans against `devices × visits-per-month` with a 95–110% tolerance band and reports direction and magnitude: `PASS`, `UNDER (n short)`, or `OVER (~Nx, expected Mx)`.

**`MANUAL_REVIEW` is a real outcome.** A document that exists but cannot be confidently attributed to the right technician or year is not the same as a missing document. Treating them as the same produces a report nobody trusts. Ambiguity gets its own verdict plus flag text naming the specific document that caused it.

**Zone-based device classification.** Device names alone are unreliable — a "Glue Board" may be an insect monitor or an interior rodent device depending on installation zone. Classification uses the zone description as a tiebreaker, with an explicit exterior designation winning over an interior keyword. Genuinely unclear cases are flagged, never guessed.

**Two levels of deduplication.** Devices are deduplicated by barcode within a work order, so a station scanned twice does not inflate the count. Work orders are deduplicated by ticket number across files, so overlapping date-range exports can be dropped in the folder together safely.

**Temp files as the communication layer.** Power Automate cannot reliably pass variables to Python as command-line arguments; special characters in addresses break it. All hand-offs go through temp files, read with an encoding-tolerant reader because Power Automate is inconsistent about UTF-8, UTF-8-with-BOM, and UTF-16.

**Heading-based table targeting.** Injected JavaScript locates tables by heading text rather than DOM index. Index-based selectors broke on every portal update; heading-based ones survived a full front-end rebuild.

**Typo normalization.** A correction dictionary fixes 47 known misspellings before any keyword matching runs.

**Selective trust between conflicting sources.** Portal expiration dates are wrong about 15% of the time. For insurance certificates, an explicit year in the title is trusted ahead of the recorded expiration date, with the date as fallback. This was a measured decision about which field lies less often.

**OCR abandoned on evidence.** Tesseract was tested for reading device counts from floor-plan PDFs and gave unreliable results across hand-drawn, scanned, and digital maps. It was replaced with one-time manual verification plus hash-based change detection, so a map is re-reviewed only when the file actually changes. The OCR code is retained in the codebase, unused, for possible future revisit.

## Reporting

`Compliance_Eval` appends rather than overwrites, so the output tab accumulates one row per location per month. Power BI reads that history and turns it into something branch managers actually receive — their own locations, their own open items, and what specifically is wrong with each — replacing the individual emails the manual process depended on. Regional and division-level views are provided as the account requires.

Access is currently handled by distributing separate reports rather than by a row-level security rule in the model. Row-level security is the correct fix and the natural next step; distribution works, but it scales with the number of branches instead of staying constant.

## Verdict meanings

| Value | Meaning |
|---|---|
| `PASS` | Found, valid, correct technician and year where applicable |
| `MISSING` | No matching document found |
| `EXPIRED` | Found but past its expiration date for the audit month |
| `WRONG_YEAR` | Found but from the prior calendar year |
| `MANUAL_REVIEW` | Found but technician, date, or title is ambiguous |
| `EXEMPT` | Does not apply — franchise location or state exception |
| `NO_DATA` | Equipment only: no work-order data for this month |

## Status

| Component | Status |
|---|---|
| Client A automation flow | Running |
| Client B automation flow | Running |
| Work-order and equipment pipeline | Running |
| Document scrape and compliance scoring | Running |
| Map change detection | Running |
| Power BI reporting layer | Running |
| Row-level security in the Power BI model | Not implemented — access handled by distribution |
| Multi-drop location automation | Deferred |

## Author

Savanna Mullins — built independently as a self-initiated project outside of a formal engineering role.
