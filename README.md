# Compliance Automation Pipeline

An automation pipeline that scrapes a web portal, evaluates 
document compliance, and generates branch-level reports.

Built to replace a fully manual Excel audit process — 
developed from scratch while learning Power Automate, 
Python, and data engineering concepts.

---

## What It Does

This pipeline automatically visits a pest control 
management portal, collects compliance documents for 
~319 service locations across two client accounts, 
evaluates whether each location meets compliance 
requirements, and outputs structured reports.

Previously this process was done entirely by hand — 
line by line in Excel, with manual emails sent to 
each branch individually.

---

## How It Works
Web Portal → Power Automate → Python → Excel → Power BI

1. **Power Automate Desktop** navigates the web portal 
   for each location, scrapes document data, and saves 
   results to temp files

2. **Python scripts** read those temp files, evaluate 
   compliance rules, and write structured output to Excel

3. **Excel workbook** stores all raw and evaluated data 
   across multiple tabs

4. **Power BI** (in progress) will deliver branch-level 
   reports with access controls via email

---

## Compliance Documents Tracked

Each location is evaluated for five document types:
- Branch Pesticide License
- Certificate of Insurance (COI)
- Technician Pesticide License
- Technician Client Certification
- IPM / cGMP Certificate

---

## Tools Used

| Tool | Purpose |
|---|---|
| Power Automate Desktop | Browser automation + web scraping |
| Python (openpyxl) | Data evaluation + Excel writes |
| JavaScript | Injected scripts for portal scraping |
| Excel | Central data storage |
| Power BI | Reporting (in progress) |

---

## Project Structure

**scripts/**
- `evaluate_compliance.py` — evaluates compliance rules and flags issues
- `write_to_excel.py` — writes evaluated data to the Excel workbook
- `move_map.py` — moves downloaded map PDFs to the correct folder
- `get_drop_number.py` — identifies the correct dropdown for multi-location addresses

**docs/**
- `project_overview.md` — plain English explanation of the full project
- `data_flow.md` — how data moves from portal to final report
- `pepsi_flow.md` — step by step documentation of the Pepsi automation flow
- `frito_flow.md` — step by step documentation of the Frito-Lay automation flow
- `excel_structure.md` — what every tab and column in the workbook means

---

## Author

Savanna Mullins
