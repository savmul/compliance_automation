# Automation Flow Documentation

## Overview

There are two Power Automate Desktop flows — one 
for Client A and one for Client B. They are 
nearly identical in structure. The only differences 
are which Excel tabs they read and which account 
they select on the portal dashboard.

The Client B flow was built by duplicating the 
Client A flow and updating three references:
- Excel tab: Client_A_Locations → Client_B_Locations
- Excel tab: Client_A_MultiDrop → Client_B_MultiDrop
- Dashboard click: Client A link → Client B link

Both flows run overnight. Do not use the computer 
while a flow is running — Power Automate controls 
Chrome and the keyboard/mouse entirely.

Total steps: ~95 per flow
All steps from 7 onward run inside a For Each loop 
— once per location address.

---

## Phase 1 — Setup (Steps 1–6)

Runs once before the location loop begins.
Loads all data into memory so Excel does not 
need to stay open during the run.

**Step 1 — Launch Excel**
Opens the compliance workbook.
Stores the connection as ExcelInstance.

**Step 2 — Set Active Worksheet**
Switches to the Locations tab for the 
current client (Client_A_Locations or 
Client_B_Locations).

**Step 3 — Read From Excel (Address List)**
Reads all addresses from Column A.
Stores as AddressList — this is what 
the loop iterates through.

**Step 4 — Set Active Worksheet (MultiDrop)**
Switches to the MultiDrop tab.
Contains addresses that require a 
non-default dropdown selection in 
the portal search results.

**Step 5 — Read From Excel (MultiDrop List)**
Reads columns A and B from the 
MultiDrop tab.
Column A = address.
Column B = dropdown number to select 
(2 or 3).
Stored as MultiDropList.

**Step 6 — Close Excel**
Closes the workbook before the loop begins.
Required — Python cannot write to Excel 
if it is already open.

---

## Phase 2 — Loop Start and Login (Steps 7–13)

Repeats for every address in AddressList.

**Step 7 — For Each (CurrentAddress in AddressList)**
Starts the main loop.
Each iteration processes one location.

**Step 8 — Set Variable (Address)**
Copies CurrentAddress into the 
Address variable for use in 
file writes later.

**Step 9 — Write Text To File (temp_address.txt)**
Writes the current address to 
C:\Automation\temp_address.txt

Power Automate cannot reliably pass 
variables to Python as command line 
arguments. Writing to a temp file is 
the workaround used throughout 
this flow.

**Step 10 — Run Application (get_drop_number.py)**
Runs the Python script that checks 
whether the current address needs 
a non-default dropdown selection.
Passes the client name as an argument 
so the script knows which MultiDrop 
tab to check.
Set to wait for completion before 
continuing — critical, or the next 
step reads an empty file.

**Step 11 — Read Text From File (temp_dropnumber.txt)**
Reads the result written by 
get_drop_number.py.
Value is 1 (default), 2, or 3.
Stored as DropNumber.

**Step 12 — Launch New Chrome**
Opens Chrome and navigates to 
the portal login page.
Stores the browser connection 
as Browser.

**Step 13 — Display Message (Manual Login)**
Pauses the flow and shows a popup 
prompting the user to log in manually
and select the correct client account.
Flow waits until the user clicks OK.

---

## Phase 3 — Navigate To Location (Steps 14–26)

**Step 14 — Click UI Element (Client link)**
Clicks the correct client account 
on the portal dashboard.

**Step 15 — Wait For Web Page Content**
Waits for the Location search bar 
to appear before continuing.

**Step 16 — Click UI Element (Search box)**
Clicks into the location search 
field to activate it.

**Step 17 — Send Keys (CurrentAddress)**
Types the current address into 
the search box.

**Step 18 — Wait For Web Page Content**
Waits for the address dropdown 
suggestion to appear after typing.

**Step 19 — Click UI Element (First result)**
Clicks the first address suggestion 
to navigate to that location.

**Steps 20–26 — Multi-Drop If/Else Block**
Checks the DropNumber value and 
selects the correct dropdown option.

- If DropNumber = 2 → JavaScript clicks 
  the second dropdown option
- Else If DropNumber = 3 → JavaScript 
  clicks the third option  
- Else → proceeds with the default 
  first result

Step 26 closes the If/Else block.

---

## Phase 4 — Map Detection (Steps 27–38)

**Step 27 — Wait For Web Page Content**
Waits for the location homepage 
to fully load.

**Step 28 — Wait 15 seconds**
Buffer for the Diagrams section 
to fully render before the 
map check runs.

**Step 29 — Run JavaScript (Map check)**
Checks the page for a Diagrams 
section containing a floor plan PDF.
Sets MapExists to YES or NO.
Also checks for "No available floor 
plans" text as a fallback.

**Step 30 — Wait 15 seconds**
Buffer after the map check before 
the result is evaluated.

**Step 31 — If MapExists = YES**
Only enters this block if a map 
was found.

**Step 32 — Run JavaScript (Format address)**
Reads and formats the current 
address for use as a filename.

**Step 33 — Run JavaScript (Click Print All)**
Finds and clicks the Print All 
button in the Diagrams section.
Triggers the map PDF download.

**Step 34 — Wait 5 seconds**
Allows the download to start.

**Step 35 — Write Text To File (temp_mapaddress.txt)**
Writes the address to temp file 
so move_map.py knows what to 
name the downloaded file.

**Step 36 — Wait 20 seconds**
Allows the PDF to fully download 
before move_map.py runs.

**Step 37 — Run Application (move_map.py)**
Moves the downloaded PDF from 
the Downloads folder to 
C:\Automation\Maps\ and renames 
it using the location address.

**Step 38 — End (closes MapExists block)**

---

## Phase 5 — Navigate To Licenses Tab (Steps 39–45)

**Step 39 — Run JavaScript (Navigate to Digital Logbook)**
Finds and clicks the Digital 
Logbook link using JavaScript 
rather than a recorded UI element.
More stable against portal UI updates.

**Step 40 — Wait For Web Page Content**
Waits for the Licenses, 
Certifications and Insurance 
tab to appear.

**Step 41 — Click UI Element (Licenses tab)**
Clicks into the Licenses, 
Certifications and Insurance tab.

**Step 42 — Wait For Web Page Content**
Waits for the Branch Contacts 
table to confirm the tab has loaded.

**Step 43 — Wait 15 seconds**
Buffer for table data to fully render.

**Step 44 — Run JavaScript (Initialize variables)**
Resets window._auditResults to 
an empty array.
Resets window._auditComplete 
to false.
Prevents data from the previous 
location carrying into this one.

**Step 45 — Run JavaScript (Table scraping setup)**
Defines helper functions used 
by all scraping steps:
- cleanText() — strips newlines 
  and extra whitespace
- getTableByHeading() — finds tables 
  by heading text rather than 
  index number

---

## Phase 6 — Scrape Four Tables (Steps 46–87)

The main data collection phase.
Four sections are scraped — each has 
its own pagination loop.

**Table 1 — Branch Contacts and Assigned 
Technicians (Steps 46–55)**
Current assigned technicians and 
their compliance documents.
Columns captured: employee name, role, 
document name, document type, 
expiration date.

**Table 2 — Other Servicing Technicians 
(Steps 56–65)**
Technicians who serviced this location 
in the past 12 months but are not 
currently assigned.
Same columns as Table 1.

**Table 3 — Insurance and Certificates 
(Steps 66–74)**
Branch-level documents: COI and 
branch license.
No employee name column.

**Table 4 — QA Inspection Reports 
(Steps 75–87)**
Located on the Report and Logs tab.
Columns captured: inspection date, 
document name.
Captures: annual assessments, 
quarterly trend reports, pesticide 
usage logs.

Each table follows the same 
pagination pattern:
Set KeepPaging variable
Loop up to 20 iterations
→ If KeepPaging = false → Exit loop
→ Wait 3 seconds
→ Run JavaScript to read next page
→ Update KeepPaging
End loop
> Note: Portal staff often upload 
> documents to whichever section is 
> convenient — there is no enforcement. 
> Python treats all four tables as one 
> unified document pool. Table source 
> is recorded but does not restrict 
> what type of document Python will 
> look for in each section.

---

## Phase 7 — Save and Write (Steps 88–92)

**Step 88 — Run JavaScript (Compile results)**
Converts window._auditResults into 
a JSON string.
Stored as DocumentData variable.

**Step 89 — Write Text To File (temp_docs.json)**
Writes DocumentData to 
C:\Automation\temp_docs.json
write_to_excel.py reads this file.

**Step 90 — Write Text To File (temp_address.txt)**
Writes CurrentAddress to 
C:\Automation\temp_address.txt
write_to_excel.py uses this to 
identify which location to write.

**Step 91 — Run Application (write_to_excel.py)**
Runs the Python script that reads 
both temp files and appends one row 
per document to Raw_Document_Pull 
in the compliance workbook.
Set to wait for completion.

**Step 92 — Close Web Browser**
Closes Chrome.
A fresh browser opens at step 12 
for the next location.

**Step 93 — End (closes the For Each loop)**

---

## Phase 8 — Post-Loop (Steps 94–95)

Runs once after all locations 
are complete.

**Step 94 — Wait 3 seconds**

**Step 95 — Run Application (evaluate_compliance.py)**
Runs the compliance evaluation 
script automatically after the 
full run completes.
Reads Raw_Document_Pull, evaluates 
all 8 checks per location, and 
writes results to Compliance_Eval.

---

## Known Constraints

**Power Fx enabled**
Power Fx mode is enabled on both flows 
and cannot be disabled without 
rebuilding from scratch. This affects 
all expression syntax:
- Numbers require = prefix: =1, =20
- String comparisons use single quotes
- Boolean logic requires Set variable 
  workarounds with If() formula
- Variables must always be inserted 
  using the {x} picker — never typed 
  manually. Typing a variable name 
  manually causes it to be written as 
  literal text instead of its value.

**UI element fragility**
Power Automate identifies page elements 
by their attributes. Portal updates 
occasionally change element IDs, 
breaking recorded clicks. Fix: 
re-record the UI element after 
an update. JavaScript-based steps 
are more resilient as they target 
elements by heading text.

**Multi-drop locations**
Approximately 17 locations return 
multiple dropdown options when 
searched. These are flagged as 
MANUAL_REVIEW in Compliance_Eval 
rather than evaluated automatically. 
A dedicated flow for multi-drop 
locations is deferred for a 
future build.
