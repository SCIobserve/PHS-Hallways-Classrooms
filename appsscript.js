// Replace your entire Apps Script with this, then Deploy > Manage deployments > Edit > New version.
//
// Sheets used:
//   - The main visits sheet (first sheet that is not "Meta", "Visits by Person", or an "Archive ..." sheet)
//   - "Visits by Person"  : running counts for the current period
//   - "Meta"              : key/value store (current period start date)
//   - "Archive YYYY-MM-DD HH.MM" : one per reset, a frozen copy of the visits from that period

const PERSON_ORDER = [
  "Kern Mojica",
  "Naima Smith",
  "Christian Zambrano",
  "Christian Cabral",
  "Austin Goldberg",
  "Carmen Vargas",
  "Eudes Budhai",
  "Ellen Hackett",
  "Janice Reid",
  "Jenna Ferris",
  "Margie Daniels",
  "Maria OlivierFlores",
  "Melanie Roman",
  "Melissa Mackhanlall"
];

const RESET_PIN = "4804";
const HEADERS = ["Timestamp", "Name", "Time", "Room", "Period"];

function isArchiveName(name) { return /^Archive /.test(name); }

function getMainSheet(ss) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const n = sheets[i].getName();
    if (n !== "Meta" && n !== "Visits by Person" && !isArchiveName(n)) return sheets[i];
  }
  return sheets[0];
}

function getMetaSheet(ss) {
  let s = ss.getSheetByName("Meta");
  if (!s) {
    s = ss.insertSheet("Meta");
    s.appendRow(["Key", "Value"]);
  }
  return s;
}

function getMeta(ss, key) {
  const rows = getMetaSheet(ss).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) if (rows[i][0] === key) return rows[i][1];
  return "";
}

function setMeta(ss, key, value) {
  const s = getMetaSheet(ss);
  const rows = s.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === key) { s.getRange(i + 1, 2).setValue(value); return; }
  }
  s.appendRow([key, value]);
}

function sheetToObjects(sheet) {
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const headers = rows[0];
  return rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[String(h).toLowerCase()] = row[i]);
    return obj;
  });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// GET
//   (no params)      -> rows of the current period
//   ?meta=1          -> { since, archives: [names, oldest first] }
//   ?sheet=NAME      -> rows of that archive sheet
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const p = (e && e.parameter) || {};

  if (p.meta) {
    const archives = ss.getSheets().map(s => s.getName()).filter(isArchiveName);
    return json({ since: getMeta(ss, "period_start") || "", archives: archives });
  }

  if (p.sheet) {
    const s = ss.getSheetByName(p.sheet);
    if (!s || !isArchiveName(p.sheet)) return json([]);
    return json(sheetToObjects(s));
  }

  return json(sheetToObjects(getMainSheet(ss)));
}

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getMainSheet(ss);
  const data = JSON.parse(e.postData.contents);

  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);

  if (data.type === "reset") {
    if (data.pin !== RESET_PIN) return ContentService.createTextOutput("BAD PIN");
    resetPeriod(ss, sheet, data.timestamp);
    return ContentService.createTextOutput("OK");
  }

  sheet.appendRow([data.timestamp, data.name, data.time, data.room, data.period]);
  updateVisitsByPerson(ss);
  return ContentService.createTextOutput("OK");
}

// Copies the current visits into a new "Archive ..." sheet, clears the main sheet, stamps the new period start.
function resetPeriod(ss, sheet, timestamp) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const now = new Date();
    const stamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH.mm");
    let name = "Archive " + stamp;
    let n = 2;
    while (ss.getSheetByName(name)) { name = "Archive " + stamp + " (" + n++ + ")"; }

    if (sheet.getLastRow() > 1) {
      const copy = sheet.copyTo(ss);
      copy.setName(name);
      copy.setTabColor("#9ca3af");
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }

    setMeta(ss, "period_start", timestamp || now.toISOString());
    updateVisitsByPerson(ss);
  } finally {
    lock.releaseLock();
  }
}

function updateVisitsByPerson(ss) {
  const rows = getMainSheet(ss).getDataRange().getValues();

  const counts = {};
  for (let i = 1; i < rows.length; i++) {
    const name = String(rows[i][1] || "").trim();
    if (name && name !== "RESET") counts[name] = (counts[name] || 0) + 1;
  }

  let vbpSheet = ss.getSheetByName("Visits by Person");
  if (!vbpSheet) vbpSheet = ss.insertSheet("Visits by Person");

  vbpSheet.clearContents();
  vbpSheet.appendRow(["Name", "Visits"]);
  PERSON_ORDER.forEach(name => vbpSheet.appendRow([name, counts[name] || 0]));
  vbpSheet.autoResizeColumns(1, 2);
}
