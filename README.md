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

## How it works

- **Dynamic Header Parsing**: The script dynamically identifies the column indices for `SECTION`, `PL NO.`, and `DRAWN` by reading the first row of each monthly sheet, ensuring robust processing even if the columns are rearranged.
- **Section Column Map**: It builds a mapping of `SECTION` -> `DRAWN` column positions from the master `PL NO.` sheet's 2-row header structure.
- **Aggregation**: It iterates through each row of the monthly sheets, sums up the drawn values, and populates the corresponding cells in the master sheet.
