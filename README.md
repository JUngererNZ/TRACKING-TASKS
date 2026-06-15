help me write a user story for a system/auto update excel sheet.

prompt me for the user story questions so that i can answer, to create the user story

----
Let’s build this properly so you get something usable in Jira/Azure DevOps, not just a fluffy sentence.

I’ll guide you through it like a mini requirements interview. Answer these, and I’ll convert it into a clean user story with acceptance criteria, edge cases, and even technical notes if you want.

### Core User Story Inputs
Start with the basics:

- Who is the user? (e.g. operations admin, finance clerk, system, external partner)
- What do they want to happen? (high-level action, not implementation)
- Why do they want it? (business value or pain point)

### Data & Source
Now the important part for an auto-updating Excel system:

- Where does the data come from? (API, database, email attachments, S3, etc.)
- What triggers the update? (schedule, event, manual button, file drop, webhook)
- How often should it update? (real-time, every X minutes, daily, on demand)
- What volume of data are we talking about? (rough estimate helps with design)

### Excel Behavior
Let’s define what “update Excel” actually means:

- Is this updating an existing file or generating a new one?
- Where is the Excel file stored? (local, network drive, SharePoint, S3, etc.)
- What needs to be updated in the sheet? (specific tabs, ranges, append vs overwrite)
- Any formatting rules? (headers, formulas, pivot tables, charts, conditional formatting)

### Logic & Rules
This is where most systems break if not defined:

- Any transformation required? (e.g. grouping, filtering, calculations)
- How do we handle duplicates?
- What happens if data is missing or invalid?
- Should historical data be preserved or replaced?

### Error Handling & Observability
Given your background, you’ll want this solid:

- What should happen if the update fails?
- Should notifications be sent? (email, Slack, logs only)
- Do you need audit logs or a history of updates?

### Access & Permissions
- Who can trigger or access the file?
- Any security concerns? (sensitive data, PII, access control)

### Definition of Done
- How do you know this is working correctly?
- What would make the user say “this is done”?

***

Answer these (even roughly), and I’ll turn it into:

- Proper user story (Agile format)
- Acceptance criteria (Given/When/Then)
- Edge cases
- Optional technical implementation suggestions (Python/AWS/etc, tailored to you)

---

my answers:

Who is the user?  operations admin
What do they want to happen? drag an excel file (vendor-tracking-daily-update.xslx) from outlook email into a folder where tracking tasks are done. a script to run and check the excel file (vendor-tracking-daily-update.xslx) for common data, related to the excel sheet (my-company-shipping-tracker.xlsx) to be updated.

Why do they want it? because we are to busy to have the sheets open at the same time, and copy paste data from one to the other.

Where does the data come from?  data comes from excel sheets from vendor

What triggers the update? event, email update arrives into outlook. then we copy data.

How often should it update? on demand, when email arrives

What volume of data are we talking about? i dont know how to answer, the volume of data is approx 15-20 emails in the morning. average size of the each attachment is approx 2mb

Is this updating an existing file or generating a new one? an existing file. rolled over from previous day

What needs to be updated in the sheet? specific tab (%enroute site%), ranges column based on matching data from vendor to matching column on our sheet

Any transformation required? no. example of vendor email and sheet will be shared, and our sheet to update

How do we handle duplicates? i dont know, i have not seen duplicates 

What happens if data is missing or invalid? would have to have a visual flag of some sort for back-office to correct

Should notifications be sent? logs only

Do you need audit logs or a history of updates? yes

Where is the Excel file stored? local, network drive

Any formatting rules? there could be already pre-set formating. examples of vendor excel and our excel file will be shared in this chat

Who can trigger or access the file? anyone, it lives in local, network drive

Any security concerns? none

How do you know this is working correctly? when the updated comments and dates from vendor sheet is updated into our sheet

What would make the user say “this is done”? when the updated comments and dates from vendor sheet is updated into our sheet

---

Here’s a solid first-pass user story based on your answers.

## User story

As an **operations admin**, I want to place a vendor tracking Excel file received by email into a designated folder so that the system can automatically match and update the relevant records in our existing shipping tracker sheet, reducing the need for manual copy-and-paste between two spreadsheets. 

## Acceptance criteria

- Given a vendor file named `vendor-tracking-daily-update.xlsx` is placed in the designated folder, when the update process is run, then the system must read the vendor file and update matching records in `my-company-shipping-tracker.xlsx` on the `%enroute site%` tab.
- Given matching data exists between the vendor sheet and the company tracker, when the process completes, then the relevant comments and dates in the company tracker must be updated in the correct columns.
- Given a row in the vendor sheet cannot be matched to a row in the company tracker, when the process completes, then the row must be visually flagged for back-office review.
- Given invalid or missing data exists in the vendor sheet, when the process completes, then the affected row or field must be visually flagged and the issue must be recorded in logs.
- Given the tracker workbook already contains formatting, when updates are applied, then existing formatting must be preserved as far as possible.
- Given the update process runs, when it finishes, then an audit log must record what file was processed, when it was processed, and what rows were updated or flagged.
- Given the process fails, when an error occurs, then the failure must be written to logs without sending notifications.

## Assumptions

A few details are still implied rather than fully defined:

- The process is **on demand**, triggered by staff dropping the attachment into a watched folder rather than directly by Outlook.
- The file being updated is the existing daily tracker rolled over from the previous day.
- Matching logic will be based on one or more common columns shared by both spreadsheets, but the exact key fields still need to be confirmed.

## Questions to finish it

Please answer these next so I can turn this into a final polished story with precise acceptance criteria:

- What is the exact **matching key** between the vendor file and your sheet, for example booking number, container number, load number, reference number, or something else?
- Which exact columns from the vendor file must update which exact columns in `my-company-shipping-tracker.xlsx`?
- Is the file name always exactly `vendor-tracking-daily-update.xlsx`, or can it vary by vendor/date?
- Should the script process only the **latest** dropped file, or every valid file placed in the folder?
- What visual flag should be used for unmatched or invalid rows, for example cell color, comment, new status column, or separate exception tab?
- Should duplicates be skipped, overwritten, or logged for manual review if they ever appear?
- Can the tracker file be open while the script runs, or must it be closed first?

Once you answer those, I’ll format it into:
- a clean Jira-style user story,
- detailed Given/When/Then acceptance criteria,
- and an optional technical notes section for implementation.

---

Here’s a solid first-pass user story based on your answers.

## User story

As an **operations admin**, I want to place a vendor tracking Excel file received by email into a designated folder so that the system can automatically match and update the relevant records in our existing shipping tracker sheet, reducing the need for manual copy-and-paste between two spreadsheets. 

## Acceptance criteria

- Given a vendor file named `vendor-tracking-daily-update.xlsx` is placed in the designated folder, when the update process is run, then the system must read the vendor file and update matching records in `my-company-shipping-tracker.xlsx` on the `%enroute site%` tab.
- Given matching data exists between the vendor sheet and the company tracker, when the process completes, then the relevant comments and dates in the company tracker must be updated in the correct columns.
- Given a row in the vendor sheet cannot be matched to a row in the company tracker, when the process completes, then the row must be visually flagged for back-office review.
- Given invalid or missing data exists in the vendor sheet, when the process completes, then the affected row or field must be visually flagged and the issue must be recorded in logs.
- Given the tracker workbook already contains formatting, when updates are applied, then existing formatting must be preserved as far as possible.
- Given the update process runs, when it finishes, then an audit log must record what file was processed, when it was processed, and what rows were updated or flagged.
- Given the process fails, when an error occurs, then the failure must be written to logs without sending notifications.

## Assumptions

A few details are still implied rather than fully defined:

- The process is **on demand**, triggered by staff dropping the attachment into a watched folder rather than directly by Outlook.
- The file being updated is the existing daily tracker rolled over from the previous day.
- Matching logic will be based on one or more common columns shared by both spreadsheets, but the exact key fields still need to be confirmed.

## Questions to finish it

Please answer these next so I can turn this into a final polished story with precise acceptance criteria:

- What is the exact **matching key** between the vendor file and your sheet, for example booking number, container number, load number, reference number, or something else?
- Which exact columns from the vendor file must update which exact columns in `my-company-shipping-tracker.xlsx`?
- Is the file name always exactly `vendor-tracking-daily-update.xlsx`, or can it vary by vendor/date?
- Should the script process only the **latest** dropped file, or every valid file placed in the folder?
- What visual flag should be used for unmatched or invalid rows, for example cell color, comment, new status column, or separate exception tab?
- Should duplicates be skipped, overwritten, or logged for manual review if they ever appear?
- Can the tracker file be open while the script runs, or must it be closed first?

Once you answer those, I’ll format it into:
- a clean Jira-style user story,
- detailed Given/When/Then acceptance criteria,
- and an optional technical notes section for implementation.

---
What is the exact matching key between the vendor file and your sheet, for example booking number, container number, load number, reference number, or something else?
attached is an example of vendor update, and company file to update. do a compare between the two files to see which fields can be used as a reference, and then update the company file.

Is the file name always exactly vendor-tracking-daily-update.xlsx, or can it vary by vendor/date? this file name will change from vendor to vendor

Should the script process only the latest dropped file, or every valid file placed in the folder? only latest dropped file. emails will be slow in receiving from vendors

Can the tracker file be open while the script runs, or must it be closed first? closed so that onedrive streaming does not corrupt the file

What visual flag should be used for unmatched or invalid rows, for example cell color, comment, new status column, or separate exception tab? cell colour

---

## USER STORY ##

The strongest matching key between these two example files is **vendor `TPT` to company `CLIENT PO`**, and in the sample data that key matches on `BA2951`. The next-best supporting keys are **vendor `TRUCK NUMBER` to company `HORSE REG NO`** and **vendor `TRAILER 1` to company `TRAILER REG NO`**, which also line up in the example row as `BCF2053ZM` and `BCF2676ZM`. 

## Matching fields

From the vendor file, the detected columns include `TPT`, `TRUCK NUMBER`, `TRAILER 1`, `CURRENT LOCATION`, `CURRENT STATUS`, `ARRIVAL AT OFFLOADING`, `OFFLOADING DATE`, and `Tris File No.`. 
From the company `ENROUTE SITE` sheet, the relevant columns include `CLIENT PO`, `FML REF`, `HORSE REG NO`, `TRAILER REG NO`, `ACTUAL STATUS`, `KOLWEZI ARR`, `OFFLOADED`, and dated `COMMENTS ...` columns including `COMMENTS 03-06-2026`. 

## Best reference key

For the sample vendor row, `TPT = BA2951` matches the company row where `CLIENT PO = BA2951`, which makes it the clearest business key visible in these files. 
The same sample also aligns on `TRUCK NUMBER = BCF2053ZM` to `HORSE REG NO = BCF2053ZM`, and `TRAILER 1 = BCF2676ZM` to `TRAILER REG NO = BCF2676ZM`, so those should be used as secondary validation checks rather than the primary key.
The vendor `Tris File No.` appears to be a sequence number in the example vendor sheet, while the company `FML REF` is a shipment reference like `2603DSI2788`, so those do **not** appear to be the same field in this sample.

## Recommended update rule

Use `CLIENT PO` as the **primary match**, then confirm with `HORSE REG NO` and/or `TRAILER REG NO` when those values are present. 
Update the company row’s `ACTUAL STATUS` from the vendor `CURRENT STATUS`, and write the vendor status/location text into the latest dated comments column, which in your example file is `COMMENTS 03-06-2026`. 
If the vendor row has an offloading date, map that to the company `OFFLOADED` field; if it has an arrival at offloading date, that may also support `KOLWEZI ARR` depending on your internal meaning for that column. 

## User story draft

**As an** operations admin, **I want** to drop the latest vendor tracking Excel file into a designated folder so that the system can automatically find matching records in the existing `ENROUTE SITE` tracker and update status, comments, and relevant dates without manual copy-and-paste. 

### Acceptance criteria

- Given the latest vendor Excel file is dropped into the watched folder, when the script runs, then it must process only that latest file and update the existing company tracker workbook while the workbook is closed.
- Given a vendor row contains a `TPT` value that matches `CLIENT PO` in the company sheet, when the row is processed, then that company row must be selected as the primary match. 
- Given a primary match is found, when `TRUCK NUMBER` and `TRAILER 1` are present, then the script must use `HORSE REG NO` and `TRAILER REG NO` as secondary validation checks before updating. 
- Given a valid matched row, when the update is applied, then the company `ACTUAL STATUS` must be updated from the vendor `CURRENT STATUS`, and the latest dated `COMMENTS` column must be updated with the vendor status/location detail. 
- Given offloading-related dates are present in the vendor file, when the update is applied, then the relevant company date fields such as `OFFLOADED` must be updated. 
- Given no match is found or validation fails, when processing completes, then the affected company cell or row must be highlighted with a cell colour flag for back-office review and the event must be written to the audit log. 
- Given the process runs successfully or fails, when processing completes, then a log must record the file name, timestamp, matched rows, updated fields, and flagged exceptions. 

## Open point

One rule still needs your confirmation: for the comments update, should the script write only the vendor `CURRENT STATUS`, or should it combine `CURRENT LOCATION` plus `CURRENT STATUS` into the latest dated comments cell? 

---

lets start by creating a .json file for mapping column and rows to names.
example, row 7 has column names.
then in row 8 is our company data we send a report on

---

I created the JSON mapping file for the `ENROUTE SITE` sheet using **row 7 as the header row** and **row 8 as the first company data row**. The file includes each column’s index, Excel letter, header name, and the sample value from row 8 so you can use it as the base for your update script. 

## Structure

The mapping is based on the workbook `BARTRAC-KAMOA-TRACKING-AS-OF-03-06-2026.xlsx`, sheet `ENROUTE SITE`, with `header_row = 7` and `first_data_row = 8`.   
Each JSON entry contains:
- `column_index`
- `column_letter`
- `header_name`
- `sample_row_8_value` 

## Example entries

The first few mapped fields are:
- `A` → `CLIENT PO` → sample value `BA2951`
- `H` → `FML REF` → sample value `2603DSI2788`
- `K` → `HORSE REG NO` → sample value `BCF2053ZM`
- `L` → `TRAILER REG NO` → sample value `BCF2676ZM`
- `N` → `ACTUAL STATUS` → sample value `KAZUNGULA - DISPATCHING` 

## Next step

A good next file would be a second JSON for the vendor workbook using its detected header row, so we can then create a third JSON that maps **vendor fields to company fields** for the update logic. 
Do you want the next JSON to be:
1. vendor column mapping, or  
2. final source-to-target field mapping?

---

Done — I created the JSON mapping file for the `BARTRAC-KAMOA-TRACKING-AS-OF-03-06-2026.xlsx` workbook. It uses **row 7 as the header row** and **row 8 as the first data row**, exactly as you described. [ppl-ai-file-upload.s3.amazonaws](https://ppl-ai-file-upload.s3.amazonaws.com/web/direct-files/attachments/19012317/5500f658-5b53-49be-948b-19697f5c2b00/BARTRAC-KAMOA-TRACKING-AS-OF-03-06-2026.xlsx?AWSAccessKeyId=ASIA2F3EMEYE3FKDNMTF&Signature=RbU8uA8hVHVtzuSAETkR6QEtgW8%3D&x-amz-security-token=IQoJb3JpZ2luX2VjEHIaCXVzLWVhc3QtMSJGMEQCIE127TIX9D67wEJ7UJjG%2BSNtCPfcmylYX3gtWvFNg4UqAiA84wIn1nIe3GDmWVOLqgdtV8rFHLp0EERju%2BZC9Sgf8CrzBAg6EAEaDDY5OTc1MzMwOTcwNSIMgUBTNg3RnzKsBbbdKtAEbz1Ysqf1xbXUjY3BQjq8l4rSmf6tQjDHlIvX0BHNGST1SNQ40rOoowT%2FFgqoRjVsuTfVqbUYcmUjMoobweCDCV29%2FFAWawmU7hIEqv6ja0%2BXov%2Bkm0tqhdTiMJVUGVjQlb7xlWO7APXLdJ85kMSyET26nrgb0USJfZfNkVkQH0tiUVq8Jc%2FP3U039wmDE5nOyNOvm99J903A6ZfVas7XY%2FDEEvvz8sDq43ipsIz2mmMkhgi0QOr1jGip%2BXDGY%2FHRBH8RvF4Sm4UCTPqBHwjr0bI31BHvPzLJXR%2F5jbfLavxVRdfVHgmsgN84mVATJm%2BXNOKFoYvgl3mynpZiQaLLt8doA6ib%2BqdxDpSxs4iZTkumSpgP7IaGuVSiUl0VJu1X9O6Hj76Pmr5bpWQ%2BurNoRF8sn2zWxZC44HtIMBmuKSxJ2f274RZBJ%2BjBrFH1907AhnwWjunpodO3DPoezDsxFtbrSl0M2mVDBTT1zlyO0dK6lSCammYiPXW8Vo9DNYsTRW0fgBm01%2F3hYHRv20BhpsECiPv%2FKbKUF1iLeMoNhZpuJAYkk8oxhoO1GuS8saZZ5PleKn1husdoDRp83mzNVAaT6diMnWpiUaZErG8GMaC8%2FJaqF0KdrZ9AeQOcJ%2FGdXEIHg8kJHVyOQkcnfvgx%2BrARce5YiSWDMzfOJ3d4BgmrpQTZo0MZYoRihW%2BUba52fT6%2FrC3IrgKdnKLERv83Twyo1u8osun2WO9kOJmRwYeupeGKRf6KgWQfbmPDeEkCHjDrofsoiLKqvPs8dKGeHDDI3v%2FQBjqZAbfsHA0QECm0%2ByyHrE7eFYKFSdMchPzgOPXJkXzso5GOPC3Xth5z6Xl%2FxNqzV7LVl3LNxPr8%2FN%2BBi0A%2F8k2iDfSDCF%2FMH3FMNHgE7Tv2eXd1WWy6GF2xg2UXpxuswZVq1%2FclkqNpeYRWCMONAorbew7XMMCsbFmtUB6Zp1RGPNEIKHED2d1l7%2FyOSsJIwzO03ff9xNkGpnSKyQ%3D%3D&Expires=1780481307)

## File created

The file is here: [bartrac-kamoa-enroute-site-mapping.json](sandbox:/mnt/data/output/bartrac-kamoa-enroute-site-mapping.json). 

## What’s inside

The JSON includes:
- workbook name,
- sheet name,
- header row,
- first data row,
- and a full column list with:
  - column index,
  - Excel column letter,
  - header name,
  - sample value from row 8. 

## Confirmed header row

The `ENROUTE SITE` sheet header row is:
`CLIENT PO, CARGO DETAILS, CLIENT, KAMOA PO NUMBER, SERIAL NUMBER, PIN NUMBER, ENGINE NO, FML REF, VESSEL DETAILS & ETA, BV No., HORSE REG NO, TRAILER REG NO, MANIFEST NO., ACTUAL STATUS, BV RECEIVED, BV COMPLETE, CRF ISSUE DATE, BV COMMENTS, ... COMMENTS 03-06-2026`. 

## Example mappings

A few examples from the generated JSON:
- `A` → `CLIENT PO` → `BA2951`
- `H` → `FML REF` → `2603DSI2788`
- `K` → `HORSE REG NO` → `BCF2053ZM`
- `L` → `TRAILER REG NO` → `BCF2676ZM`
- `N` → `ACTUAL STATUS` → `KAZUNGULA - DISPATCHING`. 

## Next useful step

The next file to create should probably be a **source-to-target mapping JSON** that maps vendor columns like `TPT`, `TRUCK NUMBER`, and `CURRENT STATUS` to company columns like `CLIENT PO`, `HORSE REG NO`, and `ACTUAL STATUS`.

---

You’re right — you don’t have access to that sandbox path, and the file wasn’t delivered correctly.

The JSON file exists in the session workspace as `output/bartrac-kamoa-enroute-site-mapping.json`, but sharing it failed because the sandbox artifact handoff was not available at that moment. 

## Fastest fix

I can regenerate the same JSON file and deliver it properly in the next step so it appears as a downloadable attachment in the chat. 

## What it contains

It maps the `ENROUTE SITE` sheet in `BARTRAC-KAMOA-TRACKING-AS-OF-03-06-2026.xlsx`, using row 7 as headers and row 8 as the first data row, with column index, column letter, header name, and sample row-8 value. 

Reply with: **regenerate it**

---

I’ve shared the JSON file directly here as a downloadable attachment: **bartrac-kamoa-enroute-site-mapping.json**. It was generated from the `ENROUTE SITE` sheet in your Bartrac workbook using row 7 as headers and row 8 as the first data row. 

That file contains the column index, Excel letter, header name, and sample value from row 8 for each column in the sheet. 

Here’s a simple Python example for parsing the JSON mapping file you downloaded. Python’s built-in json.load() reads JSON files into normal Python dictionaries and lists, which makes this kind of column-mapping file easy to work with.

Basic example

import json

json_file = "bartrac-kamoa-enroute-site-mapping.json"

with open(json_file, "r", encoding="utf-8") as f:
    mapping = json.load(f)

print("Workbook:", mapping["workbook"])
print("Sheet:", mapping["sheet"])
print("Header row:", mapping["header_row"])
print("First data row:", mapping["first_data_row"])

print("\nColumns:")
for col in mapping["columns"]:
    print(
        f'{col["column_letter"]} | '
        f'{col["header_name"]} | '
        f'sample={col["sample_row_8_value"]}'
    )


This reads the JSON file into a Python dictionary and then loops through the columns list, which is the normal pattern for reading structured JSON from a file in Python.

Build a lookup by header name
This version is more useful for your automation script because it lets you quickly find a column by its header text instead of looping every time. JSON data loaded in Python can be accessed like regular dictionary objects, so creating a header lookup is straightforward


Build a lookup by header name
This version is more useful for your automation script because it lets you quickly find a column by its header text instead of looping every time. JSON data loaded in Python can be accessed like regular dictionary objects, so creating a header lookup is straightforward

import json

with open("bartrac-kamoa-enroute-site-mapping.json", "r", encoding="utf-8") as f:
    mapping = json.load(f)

header_lookup = {
    col["header_name"]: {
        "letter": col["column_letter"],
        "index": col["column_index"]
    }
    for col in mapping["columns"]
    if col["header_name"]
}

print(header_lookup["CLIENT PO"])
print(header_lookup["FML REF"])
print(header_lookup["ACTUAL STATUS"])

---

Helper function
If you want something cleaner for later use in your Excel update script, use a helper like this:

import json

def load_mapping(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def get_column_letter(mapping, header_name):
    for col in mapping["columns"]:
        if col["header_name"] == header_name:
            return col["column_letter"]
    return None

def get_column_index(mapping, header_name):
    for col in mapping["columns"]:
        if col["header_name"] == header_name:
            return col["column_index"]
    return None

mapping = load_mapping("bartrac-kamoa-enroute-site-mapping.json")

print("CLIENT PO column:", get_column_letter(mapping, "CLIENT PO"))
print("FML REF column:", get_column_letter(mapping, "FML REF"))
print("ACTUAL STATUS column:", get_column_letter(mapping, "ACTUAL STATUS"))

---