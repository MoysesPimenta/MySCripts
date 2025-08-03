/**
 * Eurotrip Planner – AUTO w/ FX + My Maps + Charts + UI tweaks
 * v2025‑08‑03b
 * Author: ChatGPT o3
 *
 *  ✨ NEW IN THIS PATCH ✨
 *  • Alternating row colors (soft blue/white) on all main sheets.
 *  • Freeze row 1 em todas as abas (Data, Budget, Dashboard).
 *  • createMyMap() agora verifica se o serviço Drive avançado está
 *    habilitado; se não estiver, exibe instruções claras ao usuário
 *    em vez de lançar erro "Drive is not defined".
 *
 *  Como habilitar Drive avançado:
 *  ─────────────────────────────────────────────
 *  1. Apps Script ▸  ⚙️  (Project Settings) ▸ marque "Show "App sscript API requests" logs".
 *  2. Clique em **Services** (+) ▸ procure “Drive API” ▸ Add ↗︎.
 *  3. O IDE abrirá um painel pedindo para “Enable Google Drive API”
 *     no Google Cloud → clique **Enable**.
 *  4. Salve e volte à planilha; a função **createMyMap()** passará a
 *     funcionar.
 *
 *  — MENU —
 *  • ⚙️  Configurar/Atualizar Planilha
 *  • 🗺️  Gerar My Maps (KML)
 *  • 📊  Criar/Atualizar Gráficos
 *  • 📧  Lembretes de pagamento
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Planner Tools")
    .addItem("⚙️ Configurar/Atualizar Planilha", "setupSheet")
    .addItem("🗺️ Gerar My Maps (KML)", "createMyMap")
    .addItem("📊 Criar/Atualizar Gráficos", "createCharts")
    .addSeparator()
    .addItem("📧 Lembretes de pagamento (7 dias)", "sendPaymentReminders")
    .addToUi()
}

/* === Helpers === */
function categories() {
  return SpreadsheetApp.getActive()
    .getSheetByName("Lists")
    .getRange("B1:I1")
    .getValues()[0]
    .filter(String)
}
function getValidationRanges() {
  const lists = SpreadsheetApp.getActive().getSheetByName("Lists")
  return {
    categories: lists.getRange("B1:I1"),
    status: lists.getRange("B2:E2"),
    yesNo: lists.getRange("B3:C3"),
    currencies: lists.getRange("B4:C4"),
  }
}
function applyFormatting(sheet) {
  sheet.setFrozenRows(1)
  const maxR = sheet.getMaxRows()
  const maxC = sheet.getMaxColumns()
  const band = sheet
    .getRange(1, 1, maxR, maxC)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY)
  band.setHeaderRowColor("#CEE7FF") // Azul claro header
  band.setFirstBandColor("#FFFFFF") // branco
  band.setSecondBandColor("#F2F8FD") // azul muito suave
}

/* === SETUP === */
function setupSheet() {
  const ss = SpreadsheetApp.getActive()
  const data = ss.getSheetByName("Data")
  if (!data) throw new Error('Aba "Data" não encontrada.')
  const v = getValidationRanges()
  const hdr = [
    "Date",
    "Day",
    "City",
    "Description",
    "Category",
    "Status",
    "Cost Est (EUR)",
    "Cost Real (Orig)",
    "Currency",
    "FX Rate→EUR",
    "Cost Real EUR",
    "Paid (Yes/No)",
    "Booking Ref",
    "Notes",
  ]
  data.getRange(1, 1, 1, hdr.length).setValues([hdr])
  const last = Math.max(data.getLastRow(), 2)
  const rule = (r) =>
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(r, true)
      .setAllowInvalid(false)
      .build()
  data.getRange(`E2:E${last}`).setDataValidation(rule(v.categories))
  data.getRange(`F2:F${last}`).setDataValidation(rule(v.status))
  data.getRange(`I2:I${last}`).setDataValidation(rule(v.currencies))
  data.getRange(`L2:L${last}`).setDataValidation(rule(v.yesNo))
  /* formulas */
  data
    .getRange("B2")
    .setFormula('=MAP(A2:A; LAMBDA(d; IF(d=""; ""; TEXT(d; "ddd"))))')
  data
    .getRange("J2")
    .setFormula(
      '=MAP(I2:I; A2:A; LAMBDA(cur; d; IF(cur=""; ""; IF(cur="EUR";1;IFERROR(INDEX(GOOGLEFINANCE("CURRENCY:"&cur&"EUR";"price";d);2;2);"")))))'
    )
  data
    .getRange("K2")
    .setFormula(
      '=MAP(H2:H; J2:J; LAMBDA(h; fx; IF( OR(h=""; fx=""); ""; IFERROR(h*fx;""))))'
    )
  applyFormatting(data)
  buildBudgetAndDashboard()
  SpreadsheetApp.getActive().toast("Planilha configurada!", "Planner Tools")
}

/* === BUDGET & DASHBOARD === */
function buildBudgetAndDashboard() {
  const ss = SpreadsheetApp.getActive()
  const bd = ss.getSheetByName("Budget") || ss.insertSheet("Budget")
  bd.clear()
  const cats = categories()
  bd.appendRow([
    "Category",
    "Estimated (€)",
    "Real (€)",
    "Difference",
    "Paid (€)",
    "Unpaid (€)",
  ])
  cats.forEach((c, i) => {
    const r = i + 2
    bd.getRange(`A${r}`).setValue(c)
    bd.getRange(`B${r}`).setFormula(`=SUMIF(Data!E:E;A${r};Data!G:G)`)
    bd.getRange(`C${r}`).setFormula(`=SUMIF(Data!E:E;A${r};Data!K:K)`)
    bd.getRange(`D${r}`).setFormula(`=C${r}-B${r}`)
    bd.getRange(`E${r}`).setFormula(
      `=SUMIFS(Data!K:K;Data!E:E;A${r};Data!L:L;"Yes")`
    )
    bd.getRange(`F${r}`).setFormula(`=C${r}-E${r}`)
  })
  const tot = cats.length + 2
  bd.getRange(`A${tot}`).setValue("TOTAL")
  bd.getRange(`B${tot}`).setFormula(`=SUM(B2:B${tot - 1})`)
  bd.getRange(`C${tot}`).setFormula(`=SUM(C2:C${tot - 1})`)
  bd.getRange(`D${tot}`).setFormula(`=SUM(D2:D${tot - 1})`)
  bd.getRange(`E${tot}`).setFormula(`=SUM(E2:E${tot - 1})`)
  bd.getRange(`F${tot}`).setFormula(`=SUM(F2:F${tot - 1})`)
  applyFormatting(bd)
  const db = ss.getSheetByName("Dashboard") || ss.insertSheet("Dashboard")
  db.clear()
  db.appendRow(["Metric", "Value"])
  db.appendRow(["Total Estimated (€)", `=Budget!B${tot}`])
  db.appendRow(["Total Real (€)", `=Budget!C${tot}`])
  db.appendRow(["Total Paid (€)", `=Budget!E${tot}`])
  db.appendRow(["Total Unpaid (€)", `=Budget!F${tot}`])
  applyFormatting(db)
}

/* === CREATE CHARTS === (unchanged) */
function createCharts() {
  const ss = SpreadsheetApp.getActive()
  const db = ss.getSheetByName("Dashboard")
  const bd = ss.getSheetByName("Budget")
  if (!db || !bd) {
    SpreadsheetApp.getUi().alert("Execute Configurar Planilha primeiro.")
    return
  }
  db.getCharts().forEach((c) => db.removeChart(c))
  const catsCnt = categories().length
  const pie = db
    .newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(bd.getRange(`A2:C${catsCnt + 1}`))
    .setOption("title", "Gastos Reais (€) por Categoria")
    .setPosition(6, 1, 0, 0)
    .build()
  db.insertChart(pie)
  const line = db
    .newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(ss.getSheetByName("Data").getRange("A2:K"))
    .setOption("title", "Gasto acumulado (€) ao longo da viagem")
    .setPosition(6, 8, 0, 0)
    .build()
  db.insertChart(line)
  SpreadsheetApp.getActive().toast("Gráficos atualizados!", "Planner Tools")
}

/* === CREATE MY MAP === */
function createMyMap() {
  if (typeof Drive === "undefined") {
    SpreadsheetApp.getUi().alert(
      '❗ O serviço Drive avançado não está habilitado.\n\nAbra Apps Script ▸ Services (+) ▸ adicione "Drive API" e clique Enable.'
    )
    return
  }
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName("Data")
  const data = sh.getDataRange().getValues()
  const hdr = data.shift()
  const idDate = hdr.indexOf("Date"),
    idCity = hdr.indexOf("City"),
    idDesc = hdr.indexOf("Description"),
    idCost = hdr.indexOf("Cost Real EUR")
  const idLat = hdr.indexOf("Lat"),
    idLon = hdr.indexOf("Lon")
  let kml =
    '<?xml version="1.0" encoding="UTF-8"?><kml xmlns="http://www.opengis.net/kml/2.2"><Document>'
  data.forEach((r) => {
    const d = r[idDate]
    if (!(d instanceof Date)) return
    const city = r[idCity] || ""
    const desc = r[idDesc] || ""
    const cost = r[idCost] || ""
    const name =
      Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), "dd/MM") +
      " – " +
      city
    let coord = ""
    if (idLat > -1 && idLon > -1 && r[idLat] && r[idLon])
      coord = `<Point><coordinates>${r[idLon]},${r[idLat]},0</coordinates></Point>`
    kml += `<Placemark><name>${name}</name><description><![CDATA[${desc}<br/>€${cost}]]></description>${coord}</Placemark>`
  })
  kml += "</Document></kml>"
  const blob = Utilities.newBlob(
    kml,
    "application/vnd.google-earth.kml+xml",
    "eurotrip.kml"
  )
  const file = Drive.Files.insert(
    { title: "Eurotrip 2025 Map", mimeType: "application/vnd.google-apps.map" },
    blob
  )
  SpreadsheetApp.getUi().alert("Mapa criado!\n" + file.alternateLink)
}

/* === PAGAMENTOS PENDENTES (same as before) === */
function sendPaymentReminders() {
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName("Data")
  if (!sh) return
  const tz = ss.getSpreadsheetTimeZone()
  const user = Session.getActiveUser().getEmail()
  const today = new Date()
  const horizon = new Date()
  horizon.setDate(today.getDate() + 7)
  const data = sh.getDataRange().getValues()
  const hdr = data.shift()
  const idxDate = hdr.indexOf("Date")
  const idxDesc = hdr.indexOf("Description")
  const idxPaid = hdr.indexOf("Paid (Yes/No)")
  const idxReal = hdr.indexOf("Cost Real EUR")
  const unpaid = []
  data.forEach((r) => {
    const d = r[idxDate]
    if (!(d instanceof Date)) return
    if (d > horizon) return
    if (r[idxPaid] === "Yes") return
    const cost = Number(r[idxReal] || 0)
    if (cost <= 0) return
    unpaid.push({
      date: Utilities.formatDate(d, tz, "dd/MM/yyyy"),
      desc: r[idxDesc] || "—",
      cost: cost.toFixed(2),
    })
  })
  if (!unpaid.length) {
    SpreadsheetApp.getUi().alert(
      "Nenhum pagamento pendente nos próximos 7 dias."
    )
    return
  }
  const lines = unpaid.map((x) => `${x.date} • ${x.desc} • €${x.cost}`)
  MailApp.sendEmail({
    to: user,
    subject: "Lembretes de pagamento – Eurotrip Planner",
    body: lines.join("\n"),
    htmlBody:
      "<p>Pagamentos pendentes:</p><ul><li>" +
      lines.join("</li><li>") +
      "</li></ul>",
  })
  SpreadsheetApp.getUi().alert("Email de lembretes enviado!")
}
