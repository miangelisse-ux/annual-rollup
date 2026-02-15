# Annual Rollup Automation

Production-deployed PowerShell automation for aggregating and summarizing financial data across multiple Excel workbooks, generating insurance, patient, and monthly rollups.

## Production Use

This system was deployed to streamline financial reporting by consolidating multiple Excel data sources into summarized reports. It reduced manual spreadsheet work, ensured consistency, and generated ready-to-use summary files.

## Problem

Manual aggregation of Excel files for insurance, patient, and monthly financial reporting was time-consuming, error-prone, and inconsistent. This automation standardizes and speeds up the process.

## Tech Stack

- PowerShell
- Microsoft Excel COM Interop
- Windows environment

## Core Features

- Reads multiple Excel workbooks (Jan–Jul, Jul–Dec, Self Pay)
- Aggregates insurance, patient MRN, and monthly posted amounts
- Handles blank or inconsistent data
- Generates Excel summary reports
- Automatically formats numeric columns
- Auto-fits columns and bolds headers

## System Logic Overview

1. User provides paths to input Excel files and output directory.
2. Script opens Excel via COM automation (invisible, no alerts).
3. Extracts headers and normalizes them for reliable processing.
4. Aggregates:
   - Insurance rollups (Billed, Paid, Unpaid)
   - Patient MRN balances
   - Monthly posted amounts
5. Writes three summary Excel files to the output directory.
6. Closes all Excel instances safely to prevent resource leaks.

## Engineering Considerations

- COM automation used for programmatic Excel control
- Defensive error handling to avoid corrupting Excel sessions
- Dynamic header detection to handle slight variations in input files
- Auto-formatting ensures professional report output

## Future Improvements

- Add logging for each processed file
- Support more dynamic input folder structures
- Optional email delivery of summary files
- GUI wrapper for easier user interaction
- Integration with a database for long-term tracking

## Setup

1. Clone this repository:
   ```bash
   git clone https://github.com/miangelisse-ux/clinic-rollup-automation.git

2. Navigate to the Project Folder:
   ```
   cd clinic-rollup-automation

3. Place Excel Files in a convenient folder.
4. The main script is located at src/annual_rollup.ps1.

## How to Run

Open Powershell and execute:

powershell -NoProfile -ExecutionPolicy Bypass -File .\src\annual_rollup.ps1

