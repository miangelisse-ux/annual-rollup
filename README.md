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

Follow these steps to get the Annual Rollup Automation running:

1. **Requirements**
   - Windows operating system
   - Microsoft Excel installed
   - PowerShell 5.0 or higher

2. **Clone the Repository**
   ```bash
   git clone https://github.com/miangelisse-ux/annual-rollup.git
   cd annual-rollup
3. Prepare Your Input Files

Ensure you have the following Excel files ready:

   Jan–Jul data file
   Jul–Dec data file
   Self Pay data file

Place them in a folder you can easily navigate to when running the script.

4. Locate the Script

The main script is located at:
```
src/annual_rollup.ps1
```
5. Adjust PowerShell Execution Policy (if needed)

By default, PowerShell may block running scripts. The command below temporarily bypasses this restriction for the session:
```
powershell -NoProfile -ExecutionPolicy Bypass -File .\src\annual_rollup.ps1
```
6. Example File Paths

When prompted, provide full paths to your Excel files, for example:

   C:\Users\YourName\Documents\Data\Jan-Jul.xlsx
   C:\Users\YourName\Documents\Data\Jul-Dec.xlsx
   C:\Users\YourName\Documents\Data\SelfPay.xlsx

Also provide a folder path for output reports, e.g.:
```
C:\Users\YourName\Documents\RollupOutput
```
7. Run the Script

Open PowerShell and execute:
```
powershell -NoProfile -ExecutionPolicy Bypass -File .\src\annual_rollup.ps1
```
Follow the prompts to enter file paths and output folder.

Once complete, the script will generate three summary Excel files in the output folder:

   Insurance_Rollup.xlsx
   MRN_Owes_Rollup.xlsx
   Monthly_Income.xlsx
