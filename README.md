# MySCripts

## Project Purpose

This repository stores a Google Apps Script project used to automate operations in a spreadsheet. The main script resides in `USABrasil/DEP/DEP Automation Script.js` and provides utilities to highlight duplicate rows and export a subset of the data to Excel format.

## Directory Layout

```
/USABrasil
└── DEP
    ├── DEP Automation Script.js   # Google Apps Script source code
    ├── Template file.xlsx         # Spreadsheet template referenced by the script
    └── Expected Exported File.xlsx# Example of the Excel file produced by the export function
```

## Importing the Spreadsheet and Binding the Apps Script

1. Upload `USABrasil/DEP/Template file.xlsx` to your Google Drive.
2. Open the uploaded spreadsheet with Google Sheets.
3. In the sheet, select **Extensions → Apps Script** to open the Apps Script editor.
4. Replace any default code with the contents of `DEP Automation Script.js`.
5. Save the project and reload the spreadsheet.

## Running the Automation

### `exportTdsSelectSnSheetAsExcel`

1. Ensure the spreadsheet contains a sheet named `2 - TDS SELECT SNs`.
2. From the spreadsheet, run the `exportTdsSelectSnSheetAsExcel` function via **Extensions → Apps Script → Run** or create a custom menu item.
3. After execution, a sidebar appears with links to download the temporary Excel file and delete it once finished.
4. Formulas in the exported sheet are replaced with their evaluated values to avoid `#REF!` errors.

### `highlightDuplicatesDistinctColors`

1. Activate the `DEP Data` sheet in the spreadsheet.
2. Run the `highlightDuplicatesDistinctColors` function to color all duplicate values in column C using distinct colors.

### `createDepEmailDraft`

1. Ensure the spreadsheet contains a sheet named `DEP Data`.
2. Run `highlightDuplicatesDistinctColors` to open the sidebar.
3. In the sidebar, click **Create DEP Email Draft** to generate the email draft.

## Expected Exported File

The `Expected Exported File.xlsx` in `USABrasil/DEP` is a sample output produced by `exportTdsSelectSnSheetAsExcel`. When running the script, you will receive a download link to a similar file generated temporarily in your Google Drive.
