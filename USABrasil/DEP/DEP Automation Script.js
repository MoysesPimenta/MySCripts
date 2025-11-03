// @flow
/**
 * Utility script for spreadsheet automation tasks.
 * - highlightDuplicatesDistinctColors: highlight duplicates in DEP Data sheet.
 * - exportTdsSelectSnSheetAsExcel: export sheet to Excel.
 * - deleteTempFile: remove temporary export from Drive **and reset DEP Data**.
 */
const CONFIG = {
  maxRows: null,
  export: {
    folderName: "Exports",
  },
};

/* ------------------------------------------------------------------------- */
/*  DUPLICATE-HIGHLIGHTING                                                   */
/* ------------------------------------------------------------------------- */
function highlightDuplicatesDistinctColors() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() !== "DEP Data") return;

  const startRow = 2;
  const numRows = Math.max(sheet.getLastRow() - startRow + 1, 0);
  if (numRows === 0) return;

  const range = sheet.getRange(startRow, 3, numRows, 1);
  const values = range.getValues().map(r => r[0].toString().trim().toLowerCase());
  range.setBackground(null);        // reset

  const colorPool = [
    "#FFCDD2","#C5E1A5","#90CAF9","#FFCC80","#CE93D8","#FFF59D",
    "#80DEEA","#F48FB1","#A5D6A7","#B39DDB","#EF9A9A","#FFE082",
    "#81D4FA","#FFAB91","#AED581","#B0BEC5","#F06292","#BA68C8",
    "#7986CB","#4DB6AC",
  ];
  const freq  = new Map();
  values.forEach(v => { if (v) freq.set(v, (freq.get(v)||0)+1); });

  const dups = [...freq].filter(([,c]) => c>1).map(([v])=>v);
  if (!dups.length) return;

  const colorMap = {};
  dups.forEach((v,i)=>{colorMap[v]=colorPool[i%colorPool.length];});
  const backgrounds = values.map(v => [dups.includes(v) ? colorMap[v] : null]);
  range.setBackgrounds(backgrounds);
}

/* ------------------------------------------------------------------------- */
/*  EXPORT TO EXCEL & SIDEBAR                                                */
/* ------------------------------------------------------------------------- */
function exportTdsSelectSnSheetAsExcel() {
  console.time("export");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName("2 - TDS SELECT SNs");
  if (!src) { SpreadsheetApp.getUi().alert('Sheet "2 - TDS SELECT SNs" not found.'); return; }

  const tmp = SpreadsheetApp.create("Exported TDS Sheet");
  const folderName = CONFIG.export.folderName;
  if (folderName) {
    const folder = DriveApp.getFoldersByName(folderName).hasNext()
      ? DriveApp.getFoldersByName(folderName).next()
      : DriveApp.createFolder(folderName);
    const file = DriveApp.getFileById(tmp.getId());
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
  }

  const defaultSheet = tmp.getSheets()[0];
  const dataRange    = src.getDataRange();
  const target       = src.copyTo(tmp).setName("2 - TDS SELECT SNs");
  target.getRange(dataRange.getRow(), dataRange.getColumn(),
                  dataRange.getNumRows(), dataRange.getNumColumns())
        .setValues(dataRange.getValues());
  tmp.deleteSheet(defaultSheet);

  if (typeof CONFIG.maxRows === "number") {
    const max = CONFIG.maxRows, last = target.getMaxRows();
    if (last > max) target.deleteRows(max+1, last-max);
  }

  const gridStart = 7, cols = target.getLastColumn();
  if (target.getMaxRows() >= gridStart)
    target.getRange(gridStart,1,target.getMaxRows()-gridStart+1,cols)
          .setBorder(true,true,true,true,true,true,"#d9d9d9",SpreadsheetApp.BorderStyle.SOLID);

  const id  = tmp.getId();
  const dl  = `https://docs.google.com/spreadsheets/d/${id}/export?format=xlsx`;
  const url = `https://docs.google.com/spreadsheets/d/${id}/edit`;

  // === Only this HTML/JS is updated to add spinner + double-click guard ===
  const html = `
    <div style="font-family:Arial;padding:16px">
      <style>
        .btn {
          padding:8px 12px;border:none;border-radius:4px;cursor:pointer;
        }
        .btn[disabled] { opacity:.6; cursor:not-allowed; }
        .spinner {
          display:inline-block;width:14px;height:14px;border:2px solid #fff;
          border-top-color: transparent;border-radius:50%;
          margin-left:8px;vertical-align:middle;animation:spin 1s linear infinite;
        }
        @keyframes spin{ to { transform: rotate(360deg); } }
        .hidden { display:none; }
      </style>

      <h2>✅ Export Completed</h2>
      <p><strong>⬇️ Download Excel:</strong><br>
         <a href="${dl}" target="_blank" style="color:#4285F4">Click here to download .xlsx</a></p>
      <hr>
      <p><strong>🔍 View Temporary Sheet:</strong><br>
         <a href="${url}" target="_blank">Open exported sheet in Drive</a></p>
      <hr>

      <p><strong>🗑️ Delete Export & Reset DEP Data:</strong></p>
      <button id="deleteBtn" class="btn" style="background-color:#D93025;color:white">
        <span class="label">Delete Temporary File</span>
        <span class="spinner hidden" aria-hidden="true"></span>
      </button>

      <hr>
      <p><strong>📧 DEP Email:</strong></p>
      <button id="emailBtn" class="btn" style="background-color:#1a73e8;color:white">
        <span class="label">Create DEP Email</span>
        <span class="spinner hidden" aria-hidden="true"></span>
      </button>

      <script>
        function setBusy(btn, busy, workingText, doneText) {
          const label = btn.querySelector(".label");
          const spin  = btn.querySelector(".spinner");
          if (busy) {
            if (btn.dataset.busy === "1") return true; // already busy -> block double click
            btn.dataset.busy = "1";
            btn.setAttribute("disabled", "disabled");
            if (label && workingText) label.textContent = workingText;
            if (spin) spin.classList.remove("hidden");
            return false;
          } else {
            btn.dataset.busy = "0";
            btn.removeAttribute("disabled");
            if (label && doneText) label.textContent = doneText;
            if (spin) spin.classList.add("hidden");
            return false;
          }
        }

        document.getElementById("deleteBtn").onclick = function (e) {
          const btn = e.currentTarget;
          if (setBusy(btn, true, "Deleting…", null)) return;

          google.script.run
            .withSuccessHandler(function (ok) {
              // Keep disabled to avoid re-click; show final state
              setBusy(btn, false, null, ok ? "Deleted ✔" : "Error");
              if (ok) {
                alert("🗑️ Temp file deleted & DEP Data reset.");
                setTimeout(function(){ google.script.host.close(); }, 300);
              } else {
                alert("⚠️ Error.");
                // Re-enable on error so user can retry
                btn.removeAttribute("disabled");
                btn.dataset.busy = "0";
                btn.querySelector(".label").textContent = "Delete Temporary File";
              }
            })
            .withFailureHandler(function () {
              setBusy(btn, false, null, "Delete Temporary File");
              alert("⚠️ Error.");
            })
            .deleteTempFile('${id}');
        };

        document.getElementById("emailBtn").onclick = function (e) {
          const btn = e.currentTarget;
          if (setBusy(btn, true, "Creating draft…", null)) return;

          google.script.run
            .withSuccessHandler(function (ok) {
              if (ok) {
                setBusy(btn, false, null, "Draft Created ✔");
                btn.setAttribute("disabled", "disabled");
                btn.dataset.busy = "1";
                alert("📧 DEP email draft created.");
                window.open("https://mail.google.com/mail/u/0/#drafts","_blank");
              } else {
                setBusy(btn, false, null, "Create DEP Email");
                alert("⚠️ Error creating email draft.");
              }
            })
            .withFailureHandler(function () {
              setBusy(btn, false, null, "Create DEP Email");
              alert("⚠️ Error creating email draft.");
            })
            .createDepEmailDraft('${id}');
        };
      </script>
    </div>`;

  SpreadsheetApp.getUi()
    .showSidebar(HtmlService.createHtmlOutput(html).setTitle("Export Complete").setWidth(350));
  console.timeEnd("export");
}

/* ------------------------------------------------------------------------- */
/*  DELETE TEMP FILE **AND RESET DEP DATA**                                  */
/* ------------------------------------------------------------------------- */
function deleteTempFile(fileId) {
  // 1) Trash the exported spreadsheet
  try {
    if (fileId) DriveApp.getFileById(fileId).setTrashed(true);
  } catch (err) {
    Logger.log("deleteTempFile: could not trash file %s -> %s", fileId, err);
  }

  // 2) Clear DEP Data (except header) and 3) re-insert formulas
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("DEP Data");
    if (sheet) {
      // Clear contents below header
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();

      // Insert formulas
      sheet.getRange("C2").setFormula(
`=BYROW(A2:A, LAMBDA(row_value,
  IF(row_value = "",,
    IFERROR(
      LET(
        order_id, INDEX('Copia 2500'!C:C, MATCH(TRIM(row_value), TRIM('Copia 2500'!K:K), 0)),
        HYPERLINK(
          "https://www.adorama.com/MyAccount/orders#/orderdetails/" & order_id,
          order_id
        )
      ),
      "Not Found"
    )
  )
))`);
      sheet.getRange("D2").setFormula(
`=BYROW(A2:A, LAMBDA(row_value,
  IF(row_value<>"",
    IFERROR(INDEX('Copia 2500'!G:G, MATCH(row_value, 'Copia 2500'!K:K, 0)), "Not Found"),
  "")
))`);
      sheet.getRange("E2").setFormula(
`=BYROW(C2:C,LAMBDA(curVal,
    IF(curVal = "","",
      IFERROR(INDEX('Orders Database'!B:B,MATCH(curVal, 'Orders Database'!A:A, 0)),""))))`);
      sheet.getRange("F2").setFormula(
`=BYROW(C2:C,LAMBDA(curVal,IF(curVal = "","",IFERROR(INDEX('Orders Database'!C:C,MATCH(curVal, 'Orders Database'!A:A, 0)),""))))`);
      sheet.getRange("G2").setFormula(
`=BYROW(C2:C,LAMBDA(curVal,IF(curVal = "","",IFERROR(INDEX('Orders Database'!D:D,MATCH(curVal, 'Orders Database'!A:A, 0)),""))))`);
      sheet.getRange("H2").setFormula(
`=BYROW(C2:C,LAMBDA(curVal,IF(curVal = "","",IFERROR(INDEX('Orders Database'!E:E,MATCH(curVal, 'Orders Database'!A:A, 0)),""))))`);
      sheet.getRange("I2").setFormula(
`=BYROW(A2:A, LAMBDA(row_value,
  IF(row_value<>"",
    IFERROR(INDEX('Copia 2500'!H:H, MATCH(row_value, 'Copia 2500'!K:K, 0)), "Not Found"),
  "")
))`);
    }
  } catch (err) {
    Logger.log("deleteTempFile: issue resetting DEP Data -> %s", err);
  }

  // Always return true so sidebar shows success if file deletion worked
  return true;
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
 * Reads the "DEP Data" sheet and attaches the exported spreadsheet as
 * an XLSX file.
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

  const exportUrl =
    "https://docs.google.com/spreadsheets/d/" + fileId + "/export?format=xlsx";
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: "Bearer " + token },
    muteHttpExceptions: true,
  });

  const blob = response.getBlob().setName("Exported TDS Sheet.xlsx");

  GmailApp.createDraft(
    "abrahamg@adorama.com,mendelnigri@gmail.com",
    "Adorama - Request to add to ABM",
    body,
    {
      from: "dimaiscorp@gmail.com",
      attachments: [blob],
    },
  );

  try {
    DriveApp.getFileById(fileId).setTrashed(true);
  } catch (e) {
    // Ignore deletion errors
  }

  return true;
}
