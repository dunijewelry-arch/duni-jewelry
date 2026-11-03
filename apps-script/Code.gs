/**
 * Duni – reservations backend (Google Sheets + Apps Script)
 *
 * What this does:
 *  - Keeps two sheets in your spreadsheet:
 *      "פריטים"  – one row per piece: id, name, how many exist.
 *      "שריונים" – one row per reserved piece (or waitlist entry).
 *  - The website asks it "how many are left of each piece?" (doGet)
 *  - The website sends it a new reservation (doPost). It decides, piece by piece,
 *    what is reserved and what goes to the waitlist, saves it, and emails you.
 *
 * First time: run setup() once from the editor (it creates the two sheets).
 */

const OWNER_EMAIL = 'dunijewelry2@gmail.com';
const OPEN_AT = new Date('2027-02-01T00:00:00+02:00'); // sale opens, reservations close (Israel time)
const SHEET_ITEMS = 'פריטים';
const SHEET_RES = 'שריונים';
const ITEM_HEAD = ['id', 'שם', 'כמות בקולקציה', 'מחיר'];
const RES_HEAD = ['תאריך', 'קוד', 'שם', 'טלפון', 'מייל', 'id פריט', 'פריט', 'מידה', 'כמות', 'סטטוס', 'מקום בתור', 'הערות'];
const RESERVED = 'שריון';
const WAITLIST = 'רשימת המתנה';

function setup() {
  const ss = SpreadsheetApp.getActive();
  let it = ss.getSheetByName(SHEET_ITEMS);
  if (!it) {
    it = ss.insertSheet(SHEET_ITEMS);
    it.appendRow(ITEM_HEAD);
    it.setFrozenRows(1);
    it.setRightToLeft(true);
  }
  let rs = ss.getSheetByName(SHEET_RES);
  if (!rs) {
    rs = ss.insertSheet(SHEET_RES);
    rs.appendRow(RES_HEAD);
    rs.setFrozenRows(1);
    rs.setRightToLeft(true);
  }
}

/* ---------- reading ---------- */
function items_() {
  const rows = SpreadsheetApp.getActive().getSheetByName(SHEET_ITEMS).getDataRange().getValues().slice(1);
  const out = {};
  rows.forEach(r => {
    const id = String(r[0]).trim();
    if (id) out[id] = { name: String(r[1]), total: Number(r[2]) || 0, price: Number(r[3]) || 0 };
  });
  return out;
}
function counts_() {
  const rows = SpreadsheetApp.getActive().getSheetByName(SHEET_RES).getDataRange().getValues().slice(1);
  const c = {};
  rows.forEach(r => {
    const id = String(r[5]), q = Number(r[8]) || 0, st = String(r[9]);
    c[id] = c[id] || { reserved: 0, waiting: 0 };
    if (st === RESERVED) c[id].reserved += q;
    else if (st === WAITLIST) c[id].waiting += q;
  });
  return c;
}
function stock_() {
  const its = items_(), c = counts_(), out = {};
  Object.keys(its).forEach(id => {
    const r = (c[id] || {}).reserved || 0;
    out[id] = { total: its[id].total, left: Math.max(0, its[id].total - r), waiting: (c[id] || {}).waiting || 0 };
  });
  return out;
}

/* ---------- web app ---------- */
function doGet() {
  return json_({ ok: true, open: Date.now() >= OPEN_AT.getTime(), stock: stock_() });
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  if (d.website) return json_({ ok: true, code: 'X' }); // hidden spam trap filled by a bot
  if (Date.now() >= OPEN_AT.getTime()) return json_({ ok: false, error: 'closed' });

  const name = clean_(d.name, 80), phone = clean_(d.phone, 30), email = clean_(d.email, 120);
  const lines = Array.isArray(d.lines) ? d.lines.slice(0, 40) : [];
  if (!name || !/\d{9,}/.test(phone.replace(/\D/g, '')) || !lines.length || !d.consent) {
    return json_({ ok: false, error: 'missing' });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // two people at the same second can't both get the last piece
  try {
    const its = items_(), c = counts_();
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_RES);
    const code = code_();
    const now = new Date();
    const result = [];
    lines.forEach(l => {
      const id = String(l.id || '');
      const it = its[id];
      const q = Math.max(1, Math.min(20, Math.floor(Number(l.qty) || 1)));
      if (!it) return;
      c[id] = c[id] || { reserved: 0, waiting: 0 };
      const left = Math.max(0, it.total - c[id].reserved);
      const got = Math.min(q, left), wait = q - got;
      const size = clean_(l.size, 20);
      if (got) {
        sh.appendRow([now, code, name, "'" + phone, email, id, it.name, size, got, RESERVED, '', '']);
        c[id].reserved += got;
      }
      if (wait) {
        const place = c[id].waiting + 1;
        sh.appendRow([now, code, name, "'" + phone, email, id, it.name, size, wait, WAITLIST, place, '']);
        c[id].waiting += wait;
      }
      result.push({ id: id, name: it.name, size: size, reserved: got, waitlisted: wait });
    });
    if (!result.length) return json_({ ok: false, error: 'missing' });
    notify_(code, name, phone, email, result);
    return json_({ ok: true, code: code, lines: result });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- helpers ---------- */
function notify_(code, name, phone, email, lines) {
  const list = lines.map(l =>
    `• ${l.name}${l.size ? ' (' + l.size + ')' : ''}` +
    (l.reserved ? ` – שוריין ×${l.reserved}` : '') +
    (l.waitlisted ? ` – ברשימת המתנה ×${l.waitlisted}` : '')).join('\n');
  MailApp.sendEmail({
    to: OWNER_EMAIL,
    subject: `שריון חדש ${code} · ${name}`,
    body: `שם: ${name}\nטלפון: ${phone}\nמייל: ${email || '-'}\nקוד: ${code}\n\n${list}\n\nהכל נשמר בגיליון "${SHEET_RES}".`,
  });
  if (email && /@/.test(email)) {
    MailApp.sendEmail({
      to: email,
      replyTo: OWNER_EMAIL,
      name: 'Duni',
      subject: `השריון שלך ב-Duni · ${code}`,
      body: `היי ${name},\n\nתודה! קיבלנו את השריון שלך:\n\n${list}\n\n` +
        `המכירה נפתחת ב-1 בפברואר 2027 ב-00:00 (שעון ישראל). ניצור איתך קשר בוואטסאפ עם קישור אישי לרכישה.\n` +
        `התכשיטים ששוריינו שמורים לך עד 1.2 בשעה 18:00. אחר כך הם עוברים לבאות ברשימת ההמתנה.\n` +
        `מי שברשימת ההמתנה תקבל הודעה אם תכשיט משתחרר, ויהיו לה 24 שעות לרכוש אותו.\n\n` +
        `הקוד שלך: ${code}\n\nDuni ✦`,
    });
  }
}
function code_() {
  const abc = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let s = '';
  for (let i = 0; i < 6; i++) s += abc[Math.floor(Math.random() * abc.length)];
  return s;
}
function clean_(v, max) { return String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').replace(/^[=+\-@]/, "'$&").trim().slice(0, max); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
