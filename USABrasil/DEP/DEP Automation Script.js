/**
 * Generates a Gmail draft to ABM with Order ID, Machine configuration, SN, and DEP ID for each row.
 */
function generateEmailDraft() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheetByName("2 - TDS SELECT SNs")
  if (!sheet) {
    return // sheet not found
  }

  // read everything and separate header
  var data = sheet.getDataRange().getValues()
  if (data.length < 2) return // no data rows
  var headers = data.shift() // remove header row
  var idxOrder = headers.indexOf("Order ID")
  var idxCfg = headers.indexOf("Machine configuration")
  var idxSN = headers.indexOf("SN")
  var idxDep = headers.indexOf("DEP ID")

  // build body
  var body = "Hey there,\\n\\nCan you please add this machine to ABM?\\n\\n"
  data.forEach(function (row) {
    if (!row[idxSN]) return // skip empty SN rows
    body +=
      "Order ID: " +
      (row[idxOrder] || "") +
      "\\n" +
      "Machine configuration: " +
      (row[idxCfg] || "") +
      "\\n" +
      "SN: " +
      (row[idxSN] || "") +
      "\\n" +
      "ABM ID: " +
      (row[idxDep] || "") +
      "\\n\\n"
  })
  body += "Best,\\nMoyses"

  // create draft to both recipients
  GmailApp.createDraft(
    "abrahamg@adorama.com,mendelnigri@gmail.com",
    "Adorama - Request to add to ABM",
    body
  )
}
