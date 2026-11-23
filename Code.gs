/**
 * Jesus Is Lord Family Church - Visitor Check-In backend (Google Apps Script)
 * Setup: open a Google Sheet > Extensions > Apps Script > paste this file > Deploy > New deployment
 * > type "Web app" > Execute as: Me > Who has access: Anyone. Copy the Web app URL into index.html (API_URL).
 * Every check-in is stored as a row in the "Visitors" tab. Default admin PIN is 1234 (change it in the app).
 */
const TZ = "Asia/Manila";            // Davao City time (UTC+8)
const SHEET = "Visitors";
const APPROVE_BY_DEFAULT = true;     // set false to approve every visitor by hand
const HEAD = ["id", "date", "time", "name", "invitedBy", "city", "visit", "note", "approved"];
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const clean = (s, n) => String(s == null ? "" : s).trim().slice(0, n);

function sheet() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(SHEET);
  if (!s) {
    s = ss.insertSheet(SHEET);
    s.getRange("A:I").setNumberFormat("@");   // keep dates and text exactly as typed
    s.appendRow(HEAD);
    s.setFrozenRows(1);
  }
  return s;
}

function doGet() { return out({ ok: true, app: "JILFC check-in" }); }

function checkin(v) {
  const id = clean(v.id, 40), name = clean(v.name, 120), inv = clean(v.invitedBy, 120);
  if (!id || !name || !inv) return { ok: false, error: "invalid" };
  const s = sheet(), last = s.getLastRow();
  if (last > 1 && s.getRange(2, 1, last - 1, 1).getValues().some(x => String(x[0]) === id)) return { ok: true, dup: true };
  const now = new Date();
  const date = /^\d{4}-\d{2}-\d{2}$/.test(v.date) ? v.date : Utilities.formatDate(now, TZ, "yyyy-MM-dd");
  s.appendRow([id, date, Utilities.formatDate(now, TZ, "HH:mm"), name, inv, clean(v.city, 120),
    v.visit === "Second-time" ? "Second-time" : "First-time", clean(v.note, 1000), APPROVE_BY_DEFAULT ? "yes" : "no"]);
  return { ok: true };
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const r = JSON.parse(e.postData.contents);
    if (r.action === "checkin") return out(checkin(r.v || {}));

    const props = PropertiesService.getScriptProperties(), cache = CacheService.getScriptCache();
    const bad = Number(cache.get("bad") || 0);
    if (bad >= 10) return out({ ok: false, error: "locked" });                 // 5-minute lockout
    if (String(r.pin) !== (props.getProperty("PIN") || "1234")) { cache.put("bad", bad + 1, 300); return out({ ok: false, error: "pin" }); }
    cache.remove("bad");

    const s = sheet(), last = s.getLastRow();
    const all = last > 1 ? s.getRange(2, 1, last - 1, 9).getValues().map(x => x.map(String)) : [];
    const rowOf = id => { const i = all.findIndex(x => x[0] === String(id)); return i < 0 ? 0 : i + 2; };

    if (r.action === "list") {
      return out({ ok: true, rows: all.filter(x => x[1] === r.date).map(x => ({
        id: x[0], date: x[1], time: x[2], name: x[3], invitedBy: x[4], city: x[5], visit: x[6], note: x[7], approved: x[8] === "yes" })) });
    }
    if (r.action === "update") {
      const n = rowOf(r.id); if (!n) return out({ ok: false, error: "missing" });
      const col = { name: 4, invitedBy: 5, city: 6, approved: 9 }, f = r.fields || {};
      Object.keys(f).forEach(k => { if (col[k]) s.getRange(n, col[k]).setValue(k === "approved" ? (f[k] ? "yes" : "no") : clean(f[k], 120)); });
      return out({ ok: true });
    }
    if (r.action === "delete") {
      const n = rowOf(r.id); if (!n) return out({ ok: false, error: "missing" });
      s.deleteRow(n); return out({ ok: true });
    }
    if (r.action === "approveAll") {
      all.forEach((x, i) => { if (x[1] === r.date) s.getRange(i + 2, 9).setValue(r.approved ? "yes" : "no"); });
      return out({ ok: true });
    }
    if (r.action === "setPin") {
      if (!/^\d{4}$/.test(String(r.newPin))) return out({ ok: false, error: "invalid" });
      props.setProperty("PIN", String(r.newPin)); return out({ ok: true });
    }
    return out({ ok: false, error: "unknown" });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
