GEMINI CHAT

What the JSON Output Looks Like
Running this script will yield a structured configuration file containing data for each tab (e.g., ENROUTE SITE, OFFLOADED - 2026, etc.). It maps everything out cleanly:

JSON
{
    "filename": "company BARTRAC - KAMOA TRACKING AS OF 03-06-2026.xlsx",
    "sheets": {
        "ENROUTE SITE": {
            "detected_header_row": 7,
            "columns": [
                {
                    "column_index": 1,
                    "column_letter": "A",
                    "header_name": "CLIENT PO"
                },
                {
                    "column_index": 2,
                    "column_letter": "B",
                    "header_name": "CARGO DETAILS"
                },
                {
                    "column_index": 14,
                    "column_letter": "N",
                    "header_name": "ACTUAL STATUS"
                }
            ]
        }
    }
}
Key Features of this Script:
Automatic Header Detection: It automatically handles leading blank rows (like the 6 empty rows found at the top of your ENROUTE SITE tab) by analyzing cell populations.

Optimized Memory Use: Uses read_only=True which lets it inspect heavy tracking worksheets without consuming massive amounts of RAM.

Dynamic Conversion: Uses openpyxl's native helper utility to perfectly calculate double-letter cell columns (like Z matching to AA, AB, etc.).

update the script to only look at sheet "ENROUTE SITE"

---

Here is the updated Python script focused exclusively on the "ENROUTE SITE" sheet.

It targets row 7 as the header row and captures the matching sample data from row 8 to build a clean JSON configuration file, exactly as requested.

Updated Python Script
Python
import json
import os
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

def generate_enroute_site_mapping(file_path, output_json_path):
    if not os.path.exists(file_path):
        print(f"Error: The file '{file_path}' does not exist.")
        return

    print(f"Loading workbook: {file_path}...")
    # Using read_only and data_only for quick and memory-efficient reading
    wb = load_workbook(file_path, read_only=True, data_only=True)
    
    target_sheet = "ENROUTE SITE"
    if target_sheet not in wb.sheetnames:
        print(f"Error: Sheet '{target_sheet}' not found in the workbook.")
        print(f"Available sheets: {wb.sheetnames}")
        return

    ws = wb[target_sheet]
    print(f"Mapping sheet: {target_sheet}...")

    # Define exact rows based on structure: row 7 is headers, row 8 is data
    HEADER_ROW_NUM = 7
    DATA_ROW_NUM = 8

    headers = []
    sample_values = []

    # Extract row 7 and row 8 cleanly
    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if row_idx == HEADER_ROW_NUM:
            headers = list(row)
        elif row_idx == DATA_ROW_NUM:
            sample_values = list(row)
            break  # We have everything we need, exit loop early

    # Safeguard in case row 8 is shorter than row 7
    while len(sample_values) < len(headers):
        sample_values.append(None)

    sheet_columns = []

    # Build the column mapping
    for col_idx, header_value in enumerate(headers, start=1):
        # Skip completely blank columns if they trailing off the end of the sheet
        if header_value is None and (col_idx > len(sample_values) or sample_values[col_idx-1] is None):
            # If both header and sample are None, check if there are more columns left
            if all(h is None for h in headers[col_idx-1:]):
                break

        clean_header = str(header_value).strip() if header_value is not None else f"Unnamed_Column_{col_idx}"
        column_letter = get_column_letter(col_idx)
        raw_sample = sample_values[col_idx - 1]
        
        # Format dates/numbers nicely for the JSON string representation if needed
        sample_string = str(raw_sample).strip() if raw_sample is not None else ""

        sheet_columns.append({
            "column_index": col_idx,
            "column_letter": column_letter,
            "header_name": clean_header,
            "sample_row_8_value": sample_string
        })

    # Prepare final JSON layout
    mapping_output = {
        "workbook": os.path.basename(file_path),
        "sheet": target_sheet,
        "header_row": HEADER_ROW_NUM,
        "first_data_row": DATA_ROW_NUM,
        "columns": sheet_columns
    }

    # Write mapping dictionary to JSON
    with open(output_json_path, "w", encoding="utf-8") as f:
        json.dump(mapping_output, f, indent=4, ensure_ascii=False)
        
    print(f"\nSuccess! Mapping written to: {output_json_path}")

if __name__ == "__main__":
    EXCEL_FILE = "company BARTRAC - KAMOA TRACKING AS OF 03-06-2026.xlsx"
    OUTPUT_JSON = "bartrac-kamoa-enroute-site-mapping.json"
    
    generate_enroute_site_mapping(EXCEL_FILE, OUTPUT_JSON)

---

Yes, both the Python script and the resulting JSON mapping are perfectly in order. They accurately align with the actual data layout of your Excel sheet.

Here is a quick confirmation breakdown of why this looks correct:

1. Structural Accuracy
Header Alignment (row 7): The script correctly bypassed the leading 6 empty tracking rows to capture your operational headers starting at CLIENT PO on column A.

Sample Extraction (row 8): The script successfully paired every column with its true test data (e.g., Grader unit BA2951, Serial 21NK0113), ensuring your indexes match row by row.

2. Deep Column Bounds Look Right
Excel Grid Extensions: The column lettering utility correctly scaled out to double letters (CG, CH, CI, CJ, CK), capping neatly at 89 total tracked columns.

Dynamic Date Columns Captured: Looking at columns 85 to 89, it cleanly mapped your rolling timeline comments up to your latest update (COMMENTS 03-06-2026 mapping directly to Column CK).

3. Script Safeguards are Functional
The truncation logic you have in BARTRAC-KAMOA-TRACKING.py (if all(h is None for h in headers[col_idx-1:]): break) did exactly what it was designed to do: it stopped processing cleanly at column 89, preventing thousands of empty trailing Excel columns from bloating your JSON.

You are completely good to use this JSON map as the structural foundation for any automation scripts, parsing logic, or API integrations you build next!

---

```python
import pandas as pd

# Load the vendor sheets to inspect columns and data entries
oriento_2603 = pd.read_csv("example vendor ORIENTO TRACKING REPORT.xlsx - 2603DSI2788 .csv", skiprows=9)
oriento_lufu = pd.read_csv("example vendor ORIENTO TRACKING REPORT.xlsx - Lufu shipment .csv", skiprows=9)

print("2603 Columns:", oriento_2603.columns.tolist())
print("2603 Row 1 Data:", oriento_2603.iloc[0].to_dict())



```

```text
2603 Columns: ['NO', 'file', 'TPT', 'FORWARDER', 'TRUCK NUMBER', 'TRAILER 1', 'TRAILER 2', 'CAPACITY', 'REF', 'DRIVERS NAME', 'ID/PASSPORT NUMBER', 'ARRIVAL DATE', 'LOADING DATE', 'DISPATCH DATE', 'ARR GRB / BBR', 'DES GRB / BBR', 'ARR KAZ / CHIRUNDU', 'DES KAZ  / CHIRUND', 'ARRIV SAK ZAM', 'ARRIVAL SAK DRC', 'ARRIVAL WHISKEY', 'DISPATCH WHISKEY', 'ARRIVAL AT OFFLOADING', 'OFFLOADING DATE', 'CURRENT LOCATION', 'CURRENT STATUS', 'EXIT  POINT', 'ENTRY POINT', 'LOADING POINT', 'DELIVERY POINT', 'Tris File No.']
2603 Row 1 Data: {'NO': 1, 'file': 'BA2951', 'TPT': 'streamvile', 'FORWARDER': 'ORIENTO', 'TRUCK NUMBER': 'BCF2053ZM', 'TRAILER 1': 'BCF2676ZM', 'TRAILER 2': nan, 'CAPACITY': 30, 'REF': nan, 'DRIVERS NAME': 'GOODSON SICHELA', 'ID/PASSPORT NUMBER': 'ZQ265945', 'ARRIVAL DATE': '2026-05-19', 'LOADING DATE': '2026-05-19', 'DISPATCH DATE': '2026-05-20', 'ARR GRB / BBR': '2026-05-21', 'DES GRB / BBR': '2026-05-24', 'ARR KAZ / CHIRUNDU': '2026-05-30', 'DES KAZ  / CHIRUND': nan, 'ARRIV SAK ZAM': nan, 'ARRIVAL SAK DRC': nan, 'ARRIVAL WHISKEY': nan, 'DISPATCH WHISKEY': nan, 'ARRIVAL AT OFFLOADING': nan, 'OFFLOADING DATE': nan, 'CURRENT LOCATION': 'KAZUNGULA BORDER', 'CURRENT STATUS': 'TO RESUME TRANSIT ANYTIME SOON', 'EXIT  POINT': 'GRB', 'ENTRY POINT': 'KASUMBALESA', 'LOADING POINT': 'FREIGHTSTATIONS SA', 'DELIVERY POINT': 'KOLWEZI', 'Tris File No.': nan}


```

```python
# Let's verify the master file's columns/data from row 8 for matching keys
import json

with open("bartrac-kamoa-enroute-site-mapping.json", "r") as f:
    mapping = json.load(f)

# Print specific columns of interest to verify match keys
for col in mapping['columns'][:15]:
    print(f"Col {col['column_letter']}: {col['header_name']} (Sample: {col['sample_row_8_value']})")



```

```text
Col A: CLIENT PO (Sample: BA2951)
Col B: CARGO DETAILS (Sample: EG20 TIER 3 MINING GRADER)
Col C: CLIENT (Sample: BARTRAC EQUIPMENT - KAMOA COPPER)
Col D: KAMOA PO NUMBER (Sample: PO55702OX)
Col E: SERIAL NUMBER (Sample: 21NK0113)
Col F: PIN NUMBER (Sample: CAT00120VE9601161)
Col G: ENGINE NO (Sample: E7A67192)
Col H: FML REF (Sample: 2603DSI2788)
Col I: VESSEL DETAILS & ETA (Sample: VESSEL: ARIES Voy. 611N - Transhipment Vessel: CMA CGM OHIO Voy. 615S - ETA DURBAN PORT: 11-05-2026)
Col J: BV No. (Sample: COD 2025 341452)
Col K: HORSE REG NO (Sample: BCF2053ZM)
Col L: TRAILER REG NO (Sample: BCF2676ZM)
Col M: MANIFEST NO. (Sample: 2603DSI2788)
Col N: ACTUAL STATUS (Sample: KAZUNGULA - DISPATCHING)
Col O: BV RECEIVED (Sample: 10-12-2025)


```

To dynamically update Column N (`ACTUAL STATUS`) on your master tracking sheet using the vendor reports, we need deterministic **Lookup Keys** to identify the correct rows, alongside the **Target Values** to parse for the update.

The common data between your master spreadsheet (`ENROUTE SITE`) and the vendor report (`ORIENTO TRACKING REPORT`) splits cleanly into two categories:

### 1. The Unique Lookup Keys (Common Data)

These fields exist identically across both files and should be used as your compound keys to match rows with 100% certainty before running an update:

| Master Column (ENROUTE SITE) | Vendor Column (ORIENTO REPORT) | Sample Value Match | Purpose |
| --- | --- | --- | --- |
| **Column A:** `CLIENT PO` | `file` | `BA2951` | Primary Document/Identifier Key |
| **Column H:** `FML REF` | `REF` or sheet name context | `2603DSI2788` | Operational Job Reference Number |
| **Column K:** `HORSE REG NO` | `TRUCK NUMBER` | `BCF2053ZM` | Physical Logistics Asset Identifier |
| **Column L:** `TRAILER REG NO` | `TRAILER 1` | `BCF2676ZM` | Secondary Asset Verification Key |

*Note: In the vendor report, the column header for your `CLIENT PO` is named `file`, but the underlying data data matches flawlessly (`BA2951`).*

---

### 2. The Target Status Data (To Update Column N)

To map the data cleanly into your master `ACTUAL STATUS` column, the vendor report exposes two descriptive columns that can be concatenated or conditionally parsed:

* **Vendor `CURRENT LOCATION`:** `KAZUNGULA BORDER`
* **Vendor `CURRENT STATUS`:** `TO RESUME TRANSIT ANYTIME SOON`

### Next Script Step Strategy

When you update your automation pipeline, your script should search the vendor file using a combination of `file` and `TRUCK NUMBER`, pull out the `CURRENT LOCATION` and `CURRENT STATUS`, and use them to overwrite your master Column N (`ACTUAL STATUS`).

---

proceed with the next step

---

Here is the automation script that parses the vendor data from example vendor ORIENTO TRACKING REPORT.xlsx and uses the shared keys to update Column N (ACTUAL STATUS) inside the "ENROUTE SITE" sheet of your master workbook.

Script Architecture & Mapping Strategy:
Source Data Extract: It targets the tracking sheets from Oriento, standardizing their column naming conventions (e.g., mapping file to CLIENT PO and TRUCK NUMBER to HORSE REG NO).

Composite Key Matching: It builds a unique composite key combining (CLIENT PO, HORSE REG NO) for every row. This prevents any collisions if multiple drivers share similar document IDs across tracking logs.

Status Synthesis: It combines the vendor's CURRENT LOCATION and CURRENT STATUS fields into a clean status string (e.g., "KAZUNGULA BORDER - TO RESUME TRANSIT ANYTIME SOON").

Master Write-Back: It loads the master workbook, scans the active rows on "ENROUTE SITE", maps the keys, and overwrites Column N with the fresh tracking telemetry.

---

import os
import openpyxl
from openpyxl.utils import get_column_letter

def extract_vendor_updates(vendor_file_path):
    """
    Parses vendor sheets using openpyxl to build a lookup dictionary of updates.
    Composite key used: (client_po, horse_reg_no)
    """
    print(f"Extracting status updates from vendor file: {vendor_file_path}...")
    vendor_wb = openpyxl.load_workbook(vendor_file_path, read_only=True, data_only=True)
    status_updates = {}

    for sheet_name in vendor_wb.sheetnames:
        ws = vendor_wb[sheet_name]
        
        # Scan first 15 rows to locate the header block dynamically
        header_row_idx = None
        headers = []
        for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if row_idx > 15:
                break
            # Look for explicit vendor structural markers
            if "TRUCK NUMBER" in [str(c).strip() for c in row if c is not None]:
                header_row_idx = row_idx
                headers = [str(c).strip() if c is not None else "" for c in row]
                break

        if not header_row_idx:
            continue

        # Map necessary column indices (0-based for row tuples)
        try:
            # Check for alternate naming 'file' or 'CLIENT PO'
            po_col = headers.index("file") if "file" in headers else headers.index("CLIENT PO")
            truck_col = headers.index("TRUCK NUMBER")
            loc_col = headers.index("CURRENT LOCATION")
            stat_col = headers.index("CURRENT STATUS")
        except ValueError as e:
            print(f"  -> Skipping sheet '{sheet_name}': Missing required mapping column. ({e})")
            continue

        # Extract data rows following the header block
        for row_data in list(ws.iter_rows(values_only=True))[header_row_idx:]:
            # Pad short rows to match header array bounds safely
            row_cells = list(row_data) + [None] * (len(headers) - len(row_data))
            
            raw_po = row_cells[po_col]
            raw_truck = row_cells[truck_col]
            raw_loc = row_cells[loc_col]
            raw_status = row_cells[stat_col]

            if not raw_po or not raw_truck:
                continue

            # Standardize unique string keys to prevent case mismatches
            po_key = str(raw_po).strip().upper()
            truck_key = str(raw_truck).strip().upper()
            composite_key = (po_key, truck_key)

            # Consolidate spatial/operational tracking status text
            loc_str = str(raw_loc).strip() if raw_loc else ""
            stat_str = str(raw_status).strip() if raw_status else ""
            
            if loc_str and stat_str:
                combined_status = f"{loc_str} - {stat_str}"
            else:
                combined_status = loc_str or stat_str or "UNKNOWN STATUS"

            status_updates[composite_key] = combined_status

    print(f"Successfully compiled {len(status_updates)} status update paths from vendor sheets.")
    return status_updates


def update_master_sheet(master_path, updates_lookup):
    """
    Loads master tracking sheet, pairs composite rows, and writes updates into Column N.
    """
    print(f"\nOpening master workbook: {master_path}...")
    # Open without read_only mode to enable save and write capabilities
    master_wb = openpyxl.load_workbook(master_path, data_only=False)
    
    target_sheet = "ENROUTE SITE"
    if target_sheet not in master_wb.sheetnames:
        print(f"Error: Target sheet '{target_sheet}' missing from master file.")
        return

    ws = master_wb[target_sheet]
    
    # Exact coordinates extracted from structural mapping files
    HEADER_ROW = 7
    DATA_START_ROW = 8
    
    # Identify exact target column letters dynamically from header array
    headers = [str(cell.value).strip() if cell.value is not None else "" for cell in ws[HEADER_ROW]]
    
    try:
        po_col_idx = headers.index("CLIENT PO") + 1
        truck_col_idx = headers.index("HORSE REG NO") + 1
        actual_status_idx = headers.index("ACTUAL STATUS") + 1
    except ValueError as e:
        print(f"Error matching column schema positions on row {HEADER_ROW}: {e}")
        return

    updated_count = 0

    # Iterate through row records to execute targeted overrides
    for row_idx in range(DATA_START_ROW, ws.max_row + 1):
        po_val = ws.cell(row=row_idx, column=po_col_idx).value
        truck_val = ws.cell(row=row_idx, column=truck_col_idx).value

        if not po_val or not truck_val:
            continue

        po_key = str(po_val).strip().upper()
        truck_key = str(truck_val).strip().upper()
        row_composite_key = (po_key, truck_key)

        # Check if vendor matching telemetry exists for this specific equipment row
        if row_composite_key in updates_lookup:
            new_status_string = updates_lookup[row_composite_key]
            
            # Target explicit cell coordinate within Column N
            target_cell = ws.cell(row=row_idx, column=actual_status_idx)
            
            print(f"Row {row_idx:02d} [{po_key} | {truck_key}]: Overwriting '{target_cell.value}' -> '{new_status_string}'")
            
            target_cell.value = new_status_string
            updated_count += 1

    if updated_count > 0:
        master_wb.save(master_path)
        print(f"\nSuccess! Modified and saved {updated_count} status entries to master spreadsheet.")
    else:
        print("\nProcess finished. No matching composite keys were identified for update updates.")


if __name__ == "__main__":
    # Define system script paths
    MASTER_TRACKING_FILE = "company BARTRAC - KAMOA TRACKING AS OF 03-06-2026.xlsx"
    VENDOR_REPORT_FILE = "example vendor ORIENTO TRACKING REPORT.xlsx"
    
    # Run pipeline steps
    updates = extract_vendor_updates(VENDOR_REPORT_FILE)
    if updates:
        update_master_sheet(MASTER_TRACKING_FILE, updates)


---

i have 6 other company files, i can run the script against these to generate the JSON output, where by to start working on updating the company file from the vendor file


### 1. Build a Unified Schema Registry

Instead of hardcoding positions, save the JSON mapping outputs for all 6 company workbooks into a dedicated directory (e.g., `./mappings/`). Your master update script can dynamically load the correct map based on the name of the file it is processing.

```
📁 project-root/
│
├── 📁 mappings/
│   ├── bartrac_kamoa_mapping.json
│   ├── company_b_tracking_mapping.json
│   └── company_c_tracking_mapping.json
│
├── 📁 vendor_reports/
│   └── example vendor ORIENTO TRACKING REPORT.xlsx
│
└── update_pipeline.py

```

### 2. Implement the Unified Execution Pipeline

Below is the structural framework for your main processing script. It handles the ingestion of vendor data and uses your generated JSON guides to update *any* company file dynamically, regardless of its column order.

```python
import os
import json
import openpyxl

def load_json_mapping(mapping_dir, workbook_name):
    """Locates and loads the JSON structural map for a given workbook."""
    # Strip extension to look up config variations safely
    base_name = os.path.splitext(workbook_name)[0]
    for file in os.listdir(mapping_dir):
        if file.endswith(".json") and base_name in file:
            with open(os.path.join(mapping_dir, file), "r", encoding="utf-8") as f:
                return json.load(f)
    return None

def process_all_company_updates(company_files_dir, vendor_file, mapping_dir):
    """Iterates through all company files and updates them based on vendor data."""
    
    # 1. Extract fresh tracking updates from the vendor report once
    # (Using the extraction logic developed in the previous step)
    vendor_updates = extract_vendor_updates(vendor_file)
    
    if not vendor_updates:
        print("No vendor update payload found. Aborting execution.")
        return

    # 2. Process each company tracking workbook sequentially
    for filename in os.listdir(company_files_dir):
        if not filename.endswith(".xlsx") or filename.startswith("~$") or filename == vendor_file:
            continue
            
        print(f"\n" + "="*50)
        print(f"Beginning update sequence for: {filename}")
        
        # Load structural guide dynamically
        schema_guide = load_json_mapping(mapping_dir, filename)
        if not schema_guide:
            print(f"Skipping {filename}: No structural JSON mapping file found.")
            continue
            
        # 3. Execute targeted write-back using the guide
        execute_mapped_update(
            file_path=os.path.join(company_files_dir, filename),
            mapping_guide=schema_guide,
            updates_lookup=vendor_updates
        )

def execute_mapped_update(file_path, mapping_guide, updates_lookup):
    """Writes status updates to specific coordinates derived from the JSON guide."""
    wb = openpyxl.load_workbook(file_path, data_only=False)
    sheet_name = mapping_guide["sheet"]
    
    if sheet_name not in wb.sheetnames:
        print(f"Error: Target sheet '{sheet_name}' missing from workbook.")
        return
        
    ws = wb[sheet_name]
    
    # Extract structural indices safely from the JSON dictionary guide
    columns_meta = mapping_guide["columns"]
    
    def get_index(header_name):
        for col in columns_meta:
            if col["header_name"] == header_name:
                return col["column_index"]
        return None

    po_idx = get_index("CLIENT PO")
    truck_idx = get_index("HORSE REG NO")
    status_idx = get_index("ACTUAL STATUS")
    
    if not all([po_idx, truck_idx, status_idx]):
        print("Skipping workbook: Required mapping schema components not found in JSON map.")
        return

    start_row = mapping_guide["first_data_row"]
    updated_count = 0

    for row_idx in range(start_row, ws.max_row + 1):
        po_val = ws.cell(row=row_idx, column=po_idx).value
        truck_val = ws.cell(row=row_idx, column=truck_idx).value
        
        if not po_val or not truck_val:
            continue
            
        composite_key = (str(po_val).strip().upper(), str(truck_val).strip().upper())
        
        if composite_key in updates_lookup:
            ws.cell(row=row_idx, column=status_idx).value = updates_lookup[composite_key]
            updated_count += 1
            
    if updated_count > 0:
        wb.save(file_path)
        print(f"Successfully updated {updated_count} rows in '{sheet_name}'.")
    else:
        print(f"No matching composite vendor keys found for '{sheet_name}'.")

```

### Next Steps:

1. Run your mapping script over the 6 new files to populate your metadata storage.
2. Verify that each file contains exact string matches for your core tracking headers (`CLIENT PO`, `HORSE REG NO`, and `ACTUAL STATUS`). If a company file uses a variant (like `CUSTOMER PO`), simply update that specific value inside its generated `.json` file to keep the execution pipeline completely standardized.

---

