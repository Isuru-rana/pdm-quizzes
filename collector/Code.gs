// Collects PDM quiz scores into this Google Sheet.
// Setup: open the Sheet, Extensions > Apps Script, paste this file, then
// Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone.
// Put the web app URL into config.js (window.PDM_ENDPOINT).

const HEADERS = ["Time", "Name", "Score", "Total", "Percent", "Attempt", "Seconds", "Marks (Y = right)", "Picked"];

function clean(v, max) {
  let s = String(v == null ? "" : v).slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;   // stop a name being read as a formula
  return s;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const quiz = String(d.quiz || "").toLowerCase();
    if (!/^[a-z0-9-]{1,40}$/.test(quiz)) throw new Error("bad quiz id");
    const score = Number(d.score), total = Number(d.total);
    if (!(total > 0 && total <= 200 && score >= 0 && score <= total)) throw new Error("bad score");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(quiz);
    if (!sh) { sh = ss.insertSheet(quiz); }
    if (sh.getLastRow() === 0) { sh.appendRow(HEADERS); sh.setFrozenRows(1); }
    sh.appendRow([new Date(), clean(d.name, 60), score, total, Math.round(score / total * 100),
      Number(d.attempt) || 1, Number(d.seconds) || 0, clean(d.marks, 200), clean(d.picked, 200)]);
    return ContentService.createTextOutput("ok");
  } catch (err) {
    return ContentService.createTextOutput("error: " + err.message);
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput("PDM quiz collector is running.");
}
