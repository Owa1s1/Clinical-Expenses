/**
 * Clinic Expenses - Google Sheet receiver
 * Paste this whole file into Extensions > Apps Script inside the doctor's Google Sheet.
 * Change the KEY below to your own secret word, then deploy as a Web app (see HOW-TO-SET-UP.txt).
 *
 * The phone sends each entry with a unique ID. If the same entry arrives twice it is updated, never added twice.
 */
const KEY = "CHANGE-THIS-SECRET-WORD";

const HEAD = ["ID", "Date", "Month", "Group", "Category", "Paid to / taken by", "Note", "Amount", "Deleted", "Updated", "Phone"];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("Expenses");
  if (!sh) {
    sh = ss.insertSheet("Expenses");
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
    sh.getRange("A:C").setNumberFormat("@");   // keep IDs, dates and months as plain text
    sh.getRange("J:J").setNumberFormat("@");
  }
  return sh;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.key !== KEY) return out_({ ok: false, err: "wrong key" });
  if (p.op === "ping") return out_({ ok: true, pong: true });
  if (p.op !== "upsert") return out_({ ok: false, err: "unknown request" });

  let r;
  try { r = JSON.parse(p.d); } catch (x) { return out_({ ok: false, err: "bad data" }); }
  if (!r || !r.id || !r.d) return out_({ ok: false, err: "missing fields" });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const last = sh.getLastRow();
    let row = 0;
    if (last > 1) {
      const ids = sh.getRange(2, 1, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(r.id)) { row = i + 2; break; }
      }
    }
    const vals = [String(r.id), String(r.d), String(r.d).slice(0, 7), String(r.gn || ""), String(r.c || ""),
                  String(r.w || ""), String(r.n || ""), Number(r.a) || 0, r.del ? "yes" : "no", String(r.u || ""), String(r.dev || "")];
    if (row) {
      const cur = String(sh.getRange(row, 10).getValue());
      if (cur > String(r.u)) return out_({ ok: true, skipped: true });   // an older copy arrived late: ignore it
      sh.getRange(row, 1, 1, vals.length).setValues([vals]);
    } else {
      sh.getRange(last + 1, 1, 1, vals.length).setValues([vals]);
    }
    return out_({ ok: true, id: r.id });
  } finally {
    lock.releaseLock();
  }
}

/** Run this once (optional) to add a monthly summary tab. */
function setupSummary() {
  sheet_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName("Summary");
  if (!s) s = ss.insertSheet("Summary");
  s.clear();
  s.getRange("A1").setFormula(
    '=QUERY(Expenses!A2:K,"select C, D, E, sum(H) where I = \'no\' and A is not null group by C, D, E order by C desc label C \'Month\', D \'Group\', E \'Category\', sum(H) \'Total\'",0)'
  );
}
