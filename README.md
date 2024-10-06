# Excel Update Application - README

## Overview
This project automates the process of updating an Excel file by refreshing data from SAP Analysis for Office. The automation is handled by a Python script (`excel_update_application.py`) which interacts with Excel through the Windows COM interface.

The batch script (`update_excel_report.bat`) provides a simple way to run the Python script with a specific Excel file as input.

## Requirements
- Windows OS
- Python installed (preferably version 3.x)
- `pywin32` library installed (`pip install pywin32`)
- Excel installed with SAP Analysis for Office add-in

## How to Use
1. Ensure all requirements are met, especially having Python and Excel with the SAP Analysis add-in.
2. Place the following files in the same directory:
   - `excel_update_application.py` (the Python script)
   - `update_excel_report.bat` (the batch script)
3. Open a Command Prompt and navigate to the directory containing these files.
4. Run the batch file using the following command:
   ```
   update_excel_report.bat
   ```
   The batch file will execute the Python script with the Excel file `excel_to_update.xlsx` as the input.

5. The script will open Excel, refresh the SAP data, save the changes, and then close Excel.
6. Press any key to close the Command Prompt window after the process completes.

## Customizing the Input File
- To use a different Excel file, simply modify the batch file (`update_excel_report.bat`) and replace `excel_to_update.xlsx` with the name of your desired Excel file.

## Troubleshooting
- Make sure the specified Excel file exists in the same directory as the scripts.
- If Excel doesn't close properly or if an error occurs, check the Command Prompt for error messages.
- Ensure the SAP Analysis for Office add-in is properly installed and enabled in Excel.

## Notes
- The batch script pauses at the end so you can see any messages or errors that occurred during execution.

