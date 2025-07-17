// @flow
/**
 * Utility script for spreadsheet automation tasks.
 * - highlightDuplicatesDistinctColors: highlight duplicates in DEP Data sheet.
 * - exportTdsSelectSnSheetAsExcel: export sheet to Excel.
 * - deleteTempFile: remove temporary export from Drive.
 */
// Configuration object to customize export behavior.
// Set `maxRows` to limit the number of rows exported.
// Leave as `null` to export all rows.
const CONFIG = {
  maxRows: null,
  export: {
    folderName: "Exports",
  },
};

/**
 * Highlights duplicate values in column C with distinct colors.
 * Expects the active sheet to be "DEP Data".
 *
 * @returns {void}
 */
function highlightDuplicatesDistinctColors() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Silently exit if not on 'DEP Data' sheet
  if (sheet.getName() !== "DEP Data") {
    return;
  }

  const range = sheet.getRange("C3:C" + sheet.getLastRow());
  const values = range.getValues().flat();

  // Reset background color
  range.setBackground(null);

  const colorMap = {};
  const colors = [
    "#FFCDD2",
    "#C5E1A5",
    "#90CAF9",
    "#FFCC80",
    "#CE93D8",
    "#FFF59D",
    "#80DEEA",
    "#F48FB1",
    "#A5D6A7",
    "#B39DDB",
    "#EF9A9A",
    "#FFE082",
    "#81D4FA",
    "#FFAB91",
    "#AED581",
    "#B0BEC5",
    "#F06292",
    "#BA68C8",
    "#7986CB",
    "#4DB6AC",
  ];

  let colorIndex = 0;

  const duplicates = values.filter(
    (item, index) => values.indexOf(item) !== index && item !== "",
  );

  duplicates.forEach((value) => {
    if (!(value in colorMap)) {
      colorMap[value] = colors[colorIndex % colors.length];
      colorIndex++;
    }
  });

  const backgrounds = values.map((value) =>
    duplicates.includes(value) ? [colorMap[value]] : [null],
  );

  range.setBackgrounds(backgrounds);
}

/**
 * Exports the "2 - TDS SELECT SNs" sheet as a temporary Excel file.
 * Respects CONFIG.maxRows when limiting rows.
 * Opens a sidebar with download and delete links.
 *
 * @returns {void}
 */
function exportTdsSelectSnSheetAsExcel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("2 - TDS SELECT SNs");

  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert(`Sheet "2 - TDS SELECT SNs" not found.`);
    return;
  }

  const tempSpreadsheet = SpreadsheetApp.create("Exported TDS Sheet");

  const exportFolderName = CONFIG.export.folderName;
  if (exportFolderName) {
    const folders = DriveApp.getFoldersByName(exportFolderName);
    const folder = folders.hasNext()
      ? folders.next()
      : DriveApp.createFolder(exportFolderName);
    const file = DriveApp.getFileById(tempSpreadsheet.getId());
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
  }

  const tempSheet = tempSpreadsheet.getSheets()[0];
  const targetSheet = tempSheet.setName("2 - TDS SELECT SNs");

  const numCols = sourceSheet.getLastColumn();
  const maxRows =
    typeof CONFIG.maxRows === "number"
      ? CONFIG.maxRows
      : sourceSheet.getLastRow();
  const numRows = Math.min(sourceSheet.getLastRow(), maxRows);

  const sourceRange = sourceSheet.getRange(1, 1, numRows, numCols);

  // Extract values and styles
  const values = sourceRange.getValues();
  const formats = sourceRange.getNumberFormats();
  const backgrounds = sourceRange.getBackgrounds();
  const fontColors = sourceRange.getFontColors();
  const fontWeights = sourceRange.getFontWeights();
  const fontStyles = sourceRange.getFontStyles();
  const horizontalAlignments = sourceRange.getHorizontalAlignments();
  const verticalAlignments = sourceRange.getVerticalAlignments();

  // Apply values and styles
  const targetRange = targetSheet.getRange(1, 1, numRows, numCols);
  targetRange.setValues(values);
  targetRange.setNumberFormats(formats);
  targetRange.setBackgrounds(backgrounds);
  targetRange.setFontColors(fontColors);
  targetRange.setFontWeights(fontWeights);
  targetRange.setFontStyles(fontStyles);
  targetRange.setHorizontalAlignments(horizontalAlignments);
  targetRange.setVerticalAlignments(verticalAlignments);

  // Column widths
  for (let c = 1; c <= numCols; c++) {
    const width = sourceSheet.getColumnWidth(c);
    targetSheet.setColumnWidth(c, width);
  }

  // Row heights
  for (let r = 1; r <= numRows; r++) {
    const height = sourceSheet.getRowHeight(r);
    targetSheet.setRowHeight(r, height);
  }

  // Merged cells within range
  const mergedRanges = sourceSheet
    .getRange(1, 1, numRows, numCols)
    .getMergedRanges();
  mergedRanges.forEach((range) => {
    const row = range.getRow();
    const col = range.getColumn();
    const rows = range.getNumRows();
    const cols = range.getNumColumns();
    if (row + rows - 1 <= maxRows) {
      targetSheet.getRange(row, col, rows, cols).merge();
    }
  });

  // ✅ Add subtle grid-style borders from row 7 down (like template)
  const dataStartRow = 7;
  const dataRowCount = Math.max(0, numRows - dataStartRow + 1);
  if (dataRowCount > 0) {
    const dataGridRange = targetSheet.getRange(
      dataStartRow,
      1,
      dataRowCount,
      numCols,
    );
    dataGridRange.setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      "#d9d9d9",
      SpreadsheetApp.BorderStyle.SOLID,
    );
  }

  // Build sidebar with download/delete
  const fileId = tempSpreadsheet.getId();
  const downloadUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`;
  const openSheetUrl = `https://docs.google.com/spreadsheets/d/${fileId}/edit`;
  const deleteFunctionCall = `google.script.run.deleteTempFile('${fileId}')`;

  const html = `
    <div style="font-family:Arial;padding:16px">
      <h2>✅ Export Completed</h2>
      <p><strong>⬇️ Download Excel:</strong></p>
      <p><a href="${downloadUrl}" target="_blank" style="color:#4285F4">Click here to download .xlsx</a></p>
      <hr>
      <p><strong>🔍 View Temporary Sheet:</strong></p>
      <p><a href="${openSheetUrl}" target="_blank">Open exported sheet in Drive</a></p>
      <hr>
      <p><strong>🗑️ Delete Export:</strong></p>
      <button id="deleteBtn" style="background-color:#D93025;color:white;padding:8px 12px;border:none;border-radius:4px;cursor:pointer">
        Delete Temporary File
      </button>

      <script>
        document.getElementById("deleteBtn").onclick = function () {
          google.script.run
            .withSuccessHandler(function (result) {
              if (result === true) {
                alert("🗑️ Temporary file deleted successfully.");
                google.script.host.close();
              } else {
                alert("⚠️ Error deleting file.");
              }
            })
            .deleteTempFile('${fileId}');
        };
      </script>


      <!-- Optional alert popup (disabled) -->
      <!--
      <script>
        setTimeout(function() {
          alert("✅ Export ready!\\n\\nDownload: ${downloadUrl}\\nDelete: open sidebar or click delete button.");
        }, 500);
      </script>
      -->
    </div>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setTitle("Export Complete")
    .setWidth(350);

  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

/**
 * Deletes the specified Drive file.
 * @param {string} fileId ID of the file to delete.
 * @return {boolean} True if deletion succeeded, false otherwise.
 */
function deleteTempFile(fileId) {
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
    return true;
  } catch (e) {
    return false;
  }
}
