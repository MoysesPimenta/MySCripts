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

  // Build frequency map to identify duplicates without costly indexOf calls.
  const frequencyMap = new Map();
  values.forEach((value) => {
    if (value === "") {
      return;
    }
    frequencyMap.set(value, (frequencyMap.get(value) || 0) + 1);
  });

  const duplicates = Array.from(frequencyMap.entries())
    .filter(([, count]) => count > 1)
    .map(([value]) => value);

  const duplicateSet = new Set(duplicates);

  duplicates.forEach((value) => {
    if (!Object.prototype.hasOwnProperty.call(colorMap, value)) {
      colorMap[value] = colors[colorIndex % colors.length];
      colorIndex++;
    }
  });

  const backgrounds = values.map((value) =>
    duplicateSet.has(value) ? [colorMap[value]] : [null],
  );

  range.setBackgrounds(backgrounds);
}

/**
 * Exports the "2 - TDS SELECT SNs" sheet as a temporary Excel file.
 * Respects CONFIG.maxRows when limiting rows.
 * Opens a sidebar with download and delete links.
 * Formulas in the copied sheet are replaced with static values.
 *
 * @returns {void}
 */
function exportTdsSelectSnSheetAsExcel() {
  console.time("export");
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

  const defaultSheet = tempSpreadsheet.getSheets()[0];
  const dataRange = sourceSheet.getDataRange();
  const dataValues = dataRange.getValues();

  const targetSheet = sourceSheet
    .copyTo(tempSpreadsheet)
    .setName("2 - TDS SELECT SNs");

  const targetRange = targetSheet.getRange(
    dataRange.getRow(),
    dataRange.getColumn(),
    dataRange.getNumRows(),
    dataRange.getNumColumns(),
  );
  targetRange.setValues(dataValues);

  tempSpreadsheet.deleteSheet(defaultSheet);

  if (typeof CONFIG.maxRows === "number") {
    const maxRows = CONFIG.maxRows;
    const lastRow = targetSheet.getMaxRows();
    if (lastRow > maxRows) {
      targetSheet.deleteRows(maxRows + 1, lastRow - maxRows);
    }
  }

  const numCols = targetSheet.getLastColumn();

  // ✅ Add subtle grid-style borders from row 7 down (like template)
  const dataStartRow = 7;
  const totalRowCount = Math.max(
    targetSheet.getMaxRows() - dataStartRow + 1,
    0,
  );

  if (totalRowCount > 0) {
    const gridRange = targetSheet.getRange(
      dataStartRow,
      1,
      totalRowCount,
      numCols,
    );
    gridRange.setBorder(
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
      <hr>
      <p><strong>📧 DEP Email:</strong></p>
      <button id="emailBtn" style="background-color:#1a73e8;color:white;padding:8px 12px;border:none;border-radius:4px;cursor:pointer">
        Create DEP Email
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
        document.getElementById("emailBtn").onclick = function () {
          google.script.run
            .withSuccessHandler(function (result) {
              if (result === true) {
                alert("📧 DEP email draft created.");
              } else {
                alert("⚠️ Error creating email draft.");
              }
            })
            .createDepEmailDraft('${fileId}');
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

  console.timeEnd("export");
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

/**
 * Create a Gmail draft listing DEP device details.
 *
 * Reads the "DEP Data" sheet and compiles a table of devices using
 * Order ID, Machine configuration, SN, and ABM ID columns.
 * The ABM ID column can also be labeled as "DEP ID" in the sheet.
 *
 * @param {string} fileId ID of the exported file to attach.
 * @returns {boolean} True on success, false otherwise.
 */
/**
 * Find the header row that contains all required columns.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet to search.
 * @param {string[]} required Column names that must exist.
 * @param {Object<string, string[]>} synonyms Alternative names.
 * @returns {{row: number, indexes: number[]} | null} Result or null if not found.
 */
function findHeaderRow(sheet, required, synonyms) {
  const searchRows = Math.min(sheet.getLastRow(), 10);
  const data = sheet
    .getRange(1, 1, searchRows, sheet.getLastColumn())
    .getValues();

  for (let i = 0; i < data.length; i++) {
    const indexMap = {};
    data[i].forEach((h, idx) => {
      indexMap[h.toString().trim().toLowerCase()] = idx;
    });

    const indexes = required.map((name) => {
      const normalized = name.toLowerCase();
      let idx = indexMap[normalized];
      if (idx === undefined && Array.isArray(synonyms[normalized])) {
        idx = synonyms[normalized]
          .map((alt) => indexMap[alt])
          .find((j) => j !== undefined);
      }
      return idx;
    });

    if (!indexes.some((j) => j === undefined)) {
      return { row: i + 1, indexes };
    }
  }

  return null;
}

/**
 * Create a Gmail draft listing DEP device details.
 *
 * Reads the "DEP Data" sheet and attaches the exported file.
 *
 * @param {string} fileId ID of the exported file to attach.
 * @returns {boolean} True on success, false otherwise.
 */
function createDepEmailDraft(fileId) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DEP Data");

  if (!sheet) {
    return false;
  }

  const required = ["order id", "machine configuration", "sn", "abm id"];
  const synonyms = {
    sn: ["serial number", "sn"],
    "machine configuration": ["machine configuration", "machine config"],
    "abm id": ["dep id"],
  };

  const headerInfo = findHeaderRow(sheet, required, synonyms);
  if (!headerInfo) {
    SpreadsheetApp.getUi().alert(
      `Missing required columns: ${required.join(", ")}`,
    );
    return false;
  }

  const { row: headerRow, indexes } = headerInfo;
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRow) {
    return false;
  }

  const range = sheet.getRange(
    headerRow + 1,
    1,
    lastRow - headerRow,
    sheet.getLastColumn(),
  );
  const rows = range.getValues();

  const formatted = rows
    .filter((r) => r[indexes[2]])
    .map((r) => {
      const [orderId, machineConfig, sn, abmId] = indexes.map((idx) => r[idx]);
      return (
        `Order ID: ${orderId}\n` +
        `Machine configuration: ${machineConfig}\n` +
        `SN: ${sn}\n` +
        `ABM ID: ${abmId}\n`
      );
    });

  if (formatted.length === 0) {
    return false;
  }

  const body =
    "Hey there,\n\n" +
    "Can you please add this machine to ABM?\n\n" +
    formatted.join("\n") +
    "\nBest,\nMoyses";

  GmailApp.createDraft(
    "abrahamg@adorama.com,mendelnigri@gmail.com",
    "Expercom - Request to add to ABM",
    body,
    { attachments: [DriveApp.getFileById(fileId)] },
  );

  return true;
}
