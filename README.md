# RWF WSC PL Items Record Automation

This Python script automates the aggregation of drawn quantities across multiple monthly sheets in an Excel workbook and updates a master sheet with the cumulative totals.

## Overview

The script (`wsc-drawn-cumulative.py`) reads an input Excel file (`WSC-Drawn-RWF.xlsx`) that contains:
1. A master sheet named `PL NO.` with a specific header structure for sections.
2. Several monthly sheets (e.g., `APRIL 2026`, `MAY 2026`) that log drawn quantities for various `PL NO.` and `SECTION` combinations.

It calculates the cumulative sum of `DRAWN` quantities for each unique `(PL NO., SECTION)` pair across all the monthly sheets and populates the master `PL NO.` sheet with the updated totals. The final result is saved as a new file, `updated_drawn_cumulative.xlsx`.

## Prerequisites

- Python 3.x
- `openpyxl` library

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/abhi542/WSC-Drawn-RWF.git
   cd WSC-Drawn-RWF
   ```

2. Create a virtual environment (recommended):
   ```bash
   python3 -m venv venv
   source venv/bin/activate
   ```

3. Install the required dependencies:
   ```bash
   pip install openpyxl
   ```

## Usage

1. Ensure your input file is named `WSC-Drawn-RWF.xlsx` and is located in the same directory as the script.
2. Run the script:
   ```bash
   python wsc-drawn-cumulative.py
   ```
3. The script will process the data and output a new file named `updated_drawn_cumulative.xlsx` in the same directory.

## Logic Implemented

The script automates the process of tallying drawn quantities from monthly sheets into a master cumulative sheet using the following logic:

1. **Mapping Master Sheet Columns**:
   - The master sheet (`PL NO.`) has a two-tier header structure. The first row contains the `SECTION` names (e.g., EMS, EMR, EWFPS), and the second row contains sub-headers like `EAR`, `DRAWN`, and `BALANCE`.
   - The script iterates through the first two rows to build a coordinate map linking each `(SECTION, 'DRAWN')` combination to its specific column index.

2. **Parsing Monthly Sheets Dynamically**:
   - For every other sheet in the workbook (e.g., `APRIL 2026`, `MAY 2026`), the script dynamically identifies the indices for the `SECTION`, `PL NO.`, and `DRAWN` columns by reading the header row. This ensures it still works even if columns are reordered in the monthly logs.
   - It iterates through each row, extracting the PL Number, Section, and Drawn quantity, and keeps a cumulative sum in a dictionary using `(PL NO., SECTION)` as the unique key.

3. **Updating the Master Sheet**:
   - The script finally iterates through the `PL NO.` master sheet. For each valid PL Number found in the sheet, it looks up the corresponding `(PL NO., SECTION)` key in the cumulative dictionary.
   - If a match is found, it updates the specific `DRAWN` cell (identified in step 1) with the calculated cumulative total.

## Sample Data Structure

### Master Sheet (`PL NO.`)
This sheet acts as the cumulative database. It has sections spanning multiple columns.

| MPC | PL | WARD | CATEGORY | SHORT DESCRIPTION | UDM DESCRIPTION | UNIT | EAR UDM | STOCK BALANCE | EAR WSC | EMS (EAR) | EMS (DRAWN) | EMS (BALANCE) | EMR (EAR) | EMR (DRAWN) | ... |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| 1 | 75201318 | 23 | 10 | BLUE KNIT GLOVES | BLUE KNIT GLOVES | NO. | 3600 | 5087 | 3600 | 0 | 0 | 0 | 0 | 0 | ... |

*(Note: The actual sheet uses row 1 for Section names like `EMS`, `EMR`, and row 2 for subheaders like `EAR`, `DRAWN`, `BALANCE`.)*

### Monthly Sheets (e.g. `APRIL 2026`)
These sheets contain the itemized log of quantities drawn by each section.

| S.NO. | SECTION | PL NO. | ITEM SHORT DESCRIPTION | DRAWN |
| :--- | :--- | :--- | :--- | :--- |
| 1 | EMR | 40050178 | COPPER CABLE, 25SQ.MM, 4 CORE, ARMOURED | 50 |
