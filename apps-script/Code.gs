/**
 * Duni – the site's backend (Google Sheets + Apps Script)
 *
 * Your spreadsheet holds everything the site shows and collects:
 *   "פריטים"  – the jewelry on the site (edited from the admin page)
 *   "גלריה"   – gallery photos (custom pieces and stall pieces)
 *   "הגדרות"  – opening date, shipping price and the like
 *   "טקסטים"  – texts you changed on the site (the rest stay as written in texts.js)
 *   "שריונים" – every reservation and waitlist entry
 * Photos uploaded from the admin page are saved in a Drive folder called "Duni – תמונות לאתר".
 *
 * FIRST TIME (on script.google.com → New project, or from the spreadsheet: Extensions → Apps Script):
 *   1. Change ADMIN_PASSWORD below to your own password (only here, never in GitHub).
 *   2. Choose "setup" at the top and press Run. Approve the permissions.
 *   3. Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone → Deploy.
 */

const ADMIN_PASSWORD = 'CHANGE_ME';          // ← put your own password here
const OWNER_EMAIL = 'dunijewelry2@gmail.com';
const TZ = 'Asia/Jerusalem';

const SH = { items: 'פריטים', gallery: 'גלריה', settings: 'הגדרות', texts: 'טקסטים', res: 'שריונים' };

// column order of each sheet: [key, header shown in the sheet]
const ITEM_COLS = [
  ['id', 'id'], ['name', 'שם באנגלית'], ['jp', 'שם ביפנית'], ['he', 'תעתיק לעברית'], ['mean', 'משמעות'],
  ['cat', 'סוג'], ['g', 'למי'], ['price', 'מחיר'], ['qty', 'כמות'], ['sizes', 'מידות'],
  ['lead', 'תיאור קצר'], ['story', 'הסיפור'], ['made', 'איך הוא נעשה'], ['photo', 'תמונה'],
  ['visible', 'מוצג באתר'], ['added', 'נוסף בתאריך'], ['col', 'קולקציה'], ['place', 'איפה מוצג'],
];
// where a piece shows: in the shop, or in the gallery as a custom piece / a stall piece
const PLACES = ['shop', 'custom', 'stall'];
const VERSION = 2;
const GAL_COLS = [['id', 'id'], ['photo', 'תמונה'], ['title', 'כותרת'], ['kind', 'סוג'], ['added', 'נוסף בתאריך']];
const RES_HEAD = ['תאריך', 'קוד', 'שם', 'טלפון', 'מייל', 'id פריט', 'פריט', 'מידה', 'כמות', 'סטטוס', 'מקום בתור', 'הערות'];
const SETTINGS_DEFAULT = [
  ['open_at', '2027-02-01 00:00', 'פתיחת המכירה (שעון ישראל)'],
  ['hold_until', '18:00', 'עד מתי השריון שמור ביום הפתיחה'],
  ['wait_hours', '24', 'כמה שעות יש למי שברשימת ההמתנה'],
  ['ship', '25', 'מחיר משלוח בשקלים'],
  ['hero', '', 'id של התכשיט בפתיחת האתר (ריק = הראשון)'],
];
const CATS = ['rings', 'necks', 'wrists', 'ears'];
const GENDERS = ['her', 'him', 'both'];
const KINDS = ['custom', 'stall'];
const RES_STATUSES = ['שריון', 'רשימת המתנה', 'בוטל', 'נרכש'];
const RESERVED = 'שריון';
const WAITLIST = 'רשימת המתנה';

/* ================= setup ================= */
function setup() {
  const ss = ss_();
  const make = (name, head) => {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
      sh.setFrozenRows(1);
      sh.setRightToLeft(true);
      sh.getRange(2, 1, 999, head.length).setNumberFormat('@'); // keep everything as plain text
    }
    return sh;
  };
  const items = make(SH.items, ITEM_COLS.map(c => c[1]));
  make(SH.gallery, GAL_COLS.map(c => c[1]));
  const st = make(SH.settings, ['הגדרה', 'ערך', 'הסבר']);
  make(SH.texts, ['טקסט', 'תוכן']);
  const res = make(SH.res, RES_HEAD);
  res.getRange(2, 1, 999, 1).setNumberFormat('d.M.yyyy HH:mm');
  if (st.getLastRow() < 2) st.getRange(2, 1, SETTINGS_DEFAULT.length, 3).setValues(SETTINGS_DEFAULT);
  if (items.getLastRow() < 2) seedSamples_(items);
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('גיליון1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
  folder_(); // creates the photo folder and asks for Drive permission
  clearCache_();
}

/* the pieces that are on the site today, so nothing disappears on the day you connect */
function seedSamples_(sh) {
  const S = [
    ['wave', 'nami', '波', 'נאמי', 'גל', 'rings', 'her', 180, '', 'טבעת חוט דקה עם גל אחד רך בחלק העליון, כמו קו שמישהו צייר על האצבע.'],
    ['scribble', 'rakugaki', '落書き', 'רקוגאקי', 'שרבוט', 'rings', 'her', 260, '', 'טבעת חוט רחבה שמתפתלת ללולאות ולזיגזגים. שרבוט שאפשר לענוד.'],
    ['loops', 'wa', '輪', 'ווה', 'מעגל', 'rings', 'her', 240, '', 'לולאות קטנות שמחוברות בשורה לאורך חזית הטבעת.'],
    ['openring', 'musubi', '結び', 'מוסובי', 'קשר', 'rings', 'him', 220, 'מתכווננת', 'פס פתוח ששני קצותיו מתעקלים ועוברים זה ליד זה. מתכווננת.'],
    ['signet', 'yoru', '夜', 'יורו', 'לילה', 'rings', 'him', 480, '', 'פס כסף רחב עם אבן שחורה בשיבוץ כסף מוגבה.'],
    ['band', 'tsuchi', '槌', 'צוצ׳י', 'פטיש', 'rings', 'him', 190, '', 'פס כסף שטוח, פשוט ויציב, ליום-יום.'],
    ['earrings', 'koishi', '小石', 'קויישי', 'חלוק נחל', 'ears', 'her', 260, '', 'חישוקים גדולים בצורה חופשית על וו. כל אחד בצורה אחרת.'],
    ['lariat', 'shizuku', '雫', 'שיזוקו', 'טיפה', 'necks', 'her', 390, '', 'שרשרת דקה עם מעגלים שנופלים לטיפה.'],
    ['pendant', 'ishi', '石', 'אישי', 'אבן', 'necks', 'him', 290, '', 'אבן שחורה גולמית בכיפת כסף, על שרשרת דקה.'],
    ['bracelet', 'kusari', '鎖', 'קוסארי', 'שרשרת', 'wrists', 'him', 420, '', 'חוליות גדולות ולא סימטריות, מחוברות ביד לשרשרת אחת.'],
  ];
  const now = stamp_();
  const rows = S.map(r => {
    const o = { id: r[0], name: r[1], jp: r[2], he: r[3], mean: r[4], cat: r[5], g: r[6], price: r[7], qty: 5, sizes: r[8],
      lead: r[9], story: '', made: '', photo: 'assets/pieces/' + r[0] + '.webp', visible: 'כן', added: now };
    return ITEM_COLS.map(c => String(o[c[0]]));
  });
  sh.getRange(2, 1, rows.length, ITEM_COLS.length).setValues(rows);
}

/* ================= reading ================= */
// works from inside the spreadsheet (Extensions → Apps Script) or as its own project on script.google.com:
// then it uses the spreadsheet called "Duni אתר" in your Drive, or creates it
function ss_() {
  const a = SpreadsheetApp.getActive(); if (a) return a;
  const p = PropertiesService.getScriptProperties(), id = p.getProperty('SHEET_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  const it = DriveApp.getFilesByName('Duni אתר');
  while (it.hasNext()) { const f = it.next(); if (f.getMimeType() === MimeType.GOOGLE_SHEETS && !f.isTrashed()) { p.setProperty('SHEET_ID', f.getId()); return SpreadsheetApp.openById(f.getId()); } }
  const ss = SpreadsheetApp.create('Duni אתר'); p.setProperty('SHEET_ID', ss.getId()); return ss;
}
function sheet_(n) { return ss_().getSheetByName(n); }
function rows_(name, cols) {
  const sh = sheet_(name); if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, cols.length).getValues().map((r, i) => {
    const o = { _row: i + 2 }; cols.forEach((c, j) => o[c[0]] = r[j] instanceof Date ? stamp_(r[j]) : String(r[j]).trim()); return o;
  }).filter(o => o.id);
}
function ensureHeads_() {
  const sh = sheet_(SH.items); if (!sh) return;
  const head = sh.getRange(1, 1, 1, ITEM_COLS.length).getValues()[0];
  ITEM_COLS.forEach((c, i) => { if (!String(head[i] || '').trim()) sh.getRange(1, i + 1).setValue(c[1]).setFontWeight('bold'); });
}
function items_() {
  return rows_(SH.items, ITEM_COLS).map(o => Object.assign(o, {
    price: Number(o.price) || 0, qty: Math.max(0, Math.floor(Number(o.qty) || 0)), visible: !/^(לא|no|false|0)$/i.test(o.visible),
    place: PLACES.indexOf(o.place) >= 0 ? o.place : 'shop',
  }));
}
function gallery_() { return rows_(SH.gallery, GAL_COLS); }
function texts_() {
  const o = {}, sh = sheet_(SH.texts);
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => { const k = String(r[0]).trim(); if (k) o[k] = String(r[1]); });
  return o;
}
function settings_() {
  const o = {}; SETTINGS_DEFAULT.forEach(d => o[d[0]] = d[1]);
  const sh = sheet_(SH.settings);
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => {
    const k = String(r[0]).trim(); if (!k) return;
    o[k] = r[1] instanceof Date ? (k === 'open_at' ? Utilities.formatDate(r[1], TZ, 'yyyy-MM-dd HH:mm') : Utilities.formatDate(r[1], TZ, 'HH:mm')) : String(r[1]).trim();
  });
  return o;
}
function openTs_(s) {
  try { return Utilities.parseDate(s.open_at, TZ, 'yyyy-MM-dd HH:mm').getTime(); } catch (e) { return new Date('2027-02-01T00:00:00+02:00').getTime(); }
}
function resRows_() {
  const sh = sheet_(SH.res); if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, RES_HEAD.length).getValues().map((r, i) => ({
    row: i + 2, date: r[0] instanceof Date ? r[0].getTime() : String(r[0]), code: String(r[1]), name: String(r[2]), phone: String(r[3]),
    email: String(r[4]), id: String(r[5]), item: String(r[6]), size: String(r[7]), qty: Number(r[8]) || 0, status: String(r[9]),
    place: String(r[10]), notes: String(r[11]),
  })).filter(r => r.code);
}
function stock_(items, res) {
  const c = {};
  res.forEach(r => {
    c[r.id] = c[r.id] || { reserved: 0, waiting: 0, sold: 0 };
    if (r.status === RESERVED) c[r.id].reserved += r.qty;
    else if (r.status === WAITLIST) c[r.id].waiting += r.qty;
    else if (r.status === 'נרכש') c[r.id].sold += r.qty;
  });
  const out = {};
  items.forEach(it => {
    const x = c[it.id] || { reserved: 0, waiting: 0, sold: 0 };
    out[it.id] = { total: it.qty, left: Math.max(0, it.qty - x.reserved - x.sold), waiting: x.waiting, reserved: x.reserved, sold: x.sold };
  });
  return out;
}

/* ================= public: what the site reads ================= */
function doGet() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('pub');
  if (hit) return text_(hit);
  const items = items_(), s = settings_(), ts = openTs_(s);
  const pub = {
    ok: true, open: Date.now() >= ts,
    settings: { openTs: ts, hold: s.hold_until, wait: Number(s.wait_hours) || 24, ship: Number(s.ship) || 0, hero: s.hero },
    items: items.filter(i => i.visible && i.place === 'shop').map(strip_),
    gallery: items.filter(i => i.visible && i.place !== 'shop').map(i => ({ id: i.id, photo: i.photo, title: i.he || i.name, kind: i.place, col: i.col,
      name: i.name, jp: i.jp, he: i.he, cat: i.cat, lead: i.lead, story: i.story, made: i.made }))
      .concat(gallery_().map(strip_)),
    texts: texts_(),
    stock: publicStock_(stock_(items, resRows_())),
  };
  const out = JSON.stringify(pub);
  if (out.length < 95000) cache.put('pub', out, 300);
  return text_(out);
}
// visitors only see how many are left of each piece, nothing about who reserved
function publicStock_(st) { const o = {}; Object.keys(st).forEach(k => o[k] = { left: st[k].left }); return o; }
function strip_(o) { const c = Object.assign({}, o); delete c._row; delete c.added; return c; }
function clearCache_() { CacheService.getScriptCache().remove('pub'); }

/* ================= requests from the site and the admin page ================= */
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  if (d.action === 'admin') return admin_(d);
  return reserve_(d);
}

function reserve_(d) {
  if (d.website) return json_({ ok: true, code: 'X', lines: [] }); // hidden spam trap filled by a bot
  const s = settings_();
  if (Date.now() >= openTs_(s)) return json_({ ok: false, error: 'closed' });
  const name = clean_(d.name, 80), phone = clean_(d.phone, 30), email = clean_(d.email, 120);
  const lines = Array.isArray(d.lines) ? d.lines.slice(0, 40) : [];
  if (!name || !/\d{9,}/.test(phone.replace(/\D/g, '')) || !lines.length || !d.consent) return json_({ ok: false, error: 'missing' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // two people at the same second can't both get the last piece
  try {
    const items = items_(), byId = {}; items.forEach(i => byId[i.id] = i);
    const st = stock_(items, resRows_());
    const sh = sheet_(SH.res), now = new Date(), code = code_(), result = [];
    lines.forEach(l => {
      const it = byId[String(l.id || '')]; if (!it || !it.visible) return;
      const q = Math.max(1, Math.min(20, Math.floor(Number(l.qty) || 1)));
      const x = st[it.id];
      const got = Math.min(q, x.left), wait = q - got, size = clean_(l.size, 20);
      const label = it.he || it.name;
      if (got) { sh.appendRow([now, code, name, "'" + phone, email, it.id, label, size, got, RESERVED, '', '']); x.left -= got; }
      if (wait) { x.waiting += wait; sh.appendRow([now, code, name, "'" + phone, email, it.id, label, size, wait, WAITLIST, x.waiting, '']); }
      result.push({ id: it.id, name: label, size: size, reserved: got, waitlisted: wait });
    });
    if (!result.length) return json_({ ok: false, error: 'missing' });
    clearCache_();
    notify_(s, code, name, phone, email, result);
    return json_({ ok: true, code: code, lines: result });
  } finally { lock.releaseLock(); }
}

function notify_(s, code, name, phone, email, lines) {
  const list = lines.map(l => `• ${l.name}${l.size ? ' (' + l.size + ')' : ''}` +
    (l.reserved ? ` – שוריין ×${l.reserved}` : '') + (l.waitlisted ? ` – ברשימת המתנה ×${l.waitlisted}` : '')).join('\n');
  try {
    MailApp.sendEmail({ to: OWNER_EMAIL, subject: `שריון חדש ${code} · ${name}`,
      body: `שם: ${name}\nטלפון: ${phone}\nמייל: ${email || '-'}\nקוד: ${code}\n\n${list}\n\nהכל שמור בגיליון "${SH.res}" ובדף הניהול.` });
  } catch (e) {}
  if (email && /@/.test(email)) {
    const open = new Date(openTs_(s));
    const day = Utilities.formatDate(open, TZ, 'd.M.yyyy'), time = Utilities.formatDate(open, TZ, 'HH:mm');
    try {
      MailApp.sendEmail({ to: email, replyTo: OWNER_EMAIL, name: 'Duni', subject: `השריון שלך ב-Duni · ${code}`,
        body: `היי ${name},\n\nתודה! קיבלתי את השריון שלך:\n\n${list}\n\n` +
          `המכירה נפתחת ב-${day} בשעה ${time} (שעון ישראל). אשלח לך בוואטסאפ קישור אישי לרכישה.\n` +
          `התכשיטים ששוריינו שמורים לך עד ${s.hold_until} באותו יום. אחר כך הם עוברים לבאות ברשימת ההמתנה.\n` +
          `מי שברשימת ההמתנה תקבל הודעה אם תכשיט משתחרר, ויהיו לה ${s.wait_hours} שעות לרכוש אותו.\n\n` +
          `הקוד שלך: ${code}\n\nבאהבה,\nעדן ✦ Duni` });
    } catch (e) {}
  }
}

/* ================= admin ================= */
function admin_(d) {
  if (!ADMIN_PASSWORD || ADMIN_PASSWORD === 'CHANGE_ME') return json_({ ok: false, error: 'no_password' });
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('fails') || 0);
  if (fails >= 30) return json_({ ok: false, error: 'locked' });
  if (String(d.password || '') !== ADMIN_PASSWORD) { cache.put('fails', String(fails + 1), 900); return json_({ ok: false, error: 'wrong_password' }); }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const op = d.op;
    if (op === 'login' || op === 'data') return json_(adminData_());
    if (op === 'upload') return json_({ ok: true, photo: upload_(d) });
    if (op === 'saveItem') { const id = saveItem_(d.item || {}); clearCache_(); return json_(Object.assign(adminData_(), { savedId: id })); }
    if (op === 'deleteItem') { const r = deleteItem_(String(d.id || '')); if (r) return json_({ ok: false, error: r }); clearCache_(); return json_(adminData_()); }
    if (op === 'moveItem') { moveItem_(String(d.id || ''), Number(d.dir) < 0 ? -1 : 1); clearCache_(); return json_(adminData_()); }
    if (op === 'saveGallery') { saveGallery_(d.entry || {}); clearCache_(); return json_(adminData_()); }
    if (op === 'galleryToItem') { galleryToItem_(String(d.id || ''), d.item || {}); clearCache_(); return json_(adminData_()); }
    if (op === 'deleteGallery') { deleteRowById_(SH.gallery, String(d.id || ''), true); clearCache_(); return json_(adminData_()); }
    if (op === 'saveTexts') { saveTexts_(d.texts || {}); clearCache_(); return json_(adminData_()); }
    if (op === 'saveSettings') { saveSettings_(d.settings || {}); clearCache_(); return json_(adminData_()); }
    if (op === 'setRes') { setRes_(Number(d.row), String(d.status || ''), d.notes); clearCache_(); return json_(adminData_()); }
    return json_({ ok: false, error: 'bad_request' });
  } catch (err) {
    return json_({ ok: false, error: 'server', message: String(err && err.message || err) });
  } finally { lock.releaseLock(); }
}

function adminData_() {
  const items = items_(), res = resRows_();
  ensureHeads_();
  return { ok: true, version: VERSION, items: items.map(strip_), gallery: gallery_().map(strip_), settings: settings_(), texts: texts_(), reservations: res, stock: stock_(items, res) };
}

function saveItem_(it) {
  const sh = sheet_(SH.items), items = items_();
  const name = cleanText_(it.name, 60), he = cleanText_(it.he, 60);
  if (!name && !he) throw new Error('missing_name');
  let id = String(it.id || '');
  const old = items.find(x => x.id === id);
  if (!old) {
    const base = (name || 'piece').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'piece';
    id = base; let n = 2; while (items.some(x => x.id === id)) id = base + '-' + n++;
  }
  const o = {
    id: id, name: name, jp: cleanText_(it.jp, 40), he: he, mean: cleanText_(it.mean, 60),
    cat: CATS.indexOf(it.cat) >= 0 ? it.cat : 'rings', g: GENDERS.indexOf(it.g) >= 0 ? it.g : 'her',
    price: String(Math.max(0, Math.round(Number(it.price) || 0))), qty: String(Math.max(0, Math.floor(Number(it.qty) || 0))),
    sizes: cleanText_(it.sizes, 120), lead: cleanText_(it.lead, 400), story: cleanText_(it.story, 4000), made: cleanText_(it.made, 200),
    photo: cleanText_(it.photo, 300), visible: it.visible === false ? 'לא' : 'כן', added: old ? old.added : stamp_(),
    col: cleanText_(it.col, 60), place: PLACES.indexOf(it.place) >= 0 ? it.place : 'shop',
  };
  const row = ITEM_COLS.map(c => o[c[0]]);
  if (old) {
    if (old.photo && old.photo !== o.photo) trashPhoto_(old.photo);
    sh.getRange(old._row, 1, 1, row.length).setValues([row]);
  } else {
    sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setNumberFormat('@').setValues([row]);
  }
  return id;
}
function deleteItem_(id) {
  if (resRows_().some(r => r.id === id && (r.status === RESERVED || r.status === WAITLIST))) return 'has_reservations';
  deleteRowById_(SH.items, id, true);
  return '';
}
function moveItem_(id, dir) {
  const sh = sheet_(SH.items), items = items_();
  const i = items.findIndex(x => x.id === id), j = i + dir;
  if (i < 0 || j < 0 || j >= items.length) return;
  const a = sh.getRange(items[i]._row, 1, 1, ITEM_COLS.length), b = sh.getRange(items[j]._row, 1, 1, ITEM_COLS.length);
  const va = a.getValues(), vb = b.getValues(); a.setValues(vb); b.setValues(va);
}
function saveGallery_(g) {
  const sh = sheet_(SH.gallery);
  const o = { id: 'g' + Date.now().toString(36), photo: cleanText_(g.photo, 300), title: cleanText_(g.title, 80),
    kind: KINDS.indexOf(g.kind) >= 0 ? g.kind : 'custom', added: stamp_() };
  if (!o.photo) throw new Error('missing_photo');
  sh.insertRowBefore(2); // newest first
  sh.getRange(2, 1, 1, GAL_COLS.length).setNumberFormat('@').setValues([GAL_COLS.map(c => o[c[0]])]);
}
function galleryToItem_(id, it) {
  const g = gallery_().find(x => x.id === id); if (!g) throw new Error('not_found');
  saveItem_(Object.assign({ he: g.title, name: '', photo: g.photo, visible: false }, it, { id: '' }));
  deleteRowById_(SH.gallery, id, false);
}
function deleteRowById_(name, id, trash) {
  const cols = name === SH.items ? ITEM_COLS : GAL_COLS;
  const r = rows_(name, cols).find(x => x.id === id); if (!r) return;
  if (trash && r.photo) trashPhoto_(r.photo);
  sheet_(name).deleteRow(r._row);
}
function saveSettings_(s) {
  const sh = sheet_(SH.settings);
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
  const ok = {
    open_at: v => /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(v), hold_until: v => /^\d{2}:\d{2}$/.test(v),
    wait_hours: v => /^\d{1,3}$/.test(v), ship: v => /^\d{1,4}$/.test(v), hero: v => /^[a-z0-9-]{0,60}$/.test(v),
  };
  Object.keys(ok).forEach(k => {
    if (!(k in s)) return; const v = String(s[k]).trim(); if (!ok[k](v)) return;
    const i = rows.indexOf(k);
    if (i >= 0) sh.getRange(i + 2, 2).setNumberFormat('@').setValue(v);
    else { const d = SETTINGS_DEFAULT.find(x => x[0] === k); sh.appendRow([k, "'" + v, d ? d[2] : '']); rows.push(k); }
  });
}
// texts: {key: new text}; an empty text goes back to the default
function saveTexts_(t) {
  const sh = sheet_(SH.texts);
  const keys = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
  Object.keys(t).forEach(k => {
    if (!/^[a-z0-9_]{1,40}$/.test(k)) return;
    const v = cleanText_(t[k], 20000), i = keys.indexOf(k);
    if (i >= 0) sh.getRange(i + 2, 2).setNumberFormat('@').setValue(v);
    else if (v) { sh.appendRow([k, '']); sh.getRange(sh.getLastRow(), 2).setNumberFormat('@').setValue(v); keys.push(k); }
  });
}
function setRes_(row, status, notes) {
  const sh = sheet_(SH.res);
  if (!(row >= 2 && row <= sh.getLastRow())) throw new Error('bad_row');
  if (RES_STATUSES.indexOf(status) >= 0) sh.getRange(row, 10).setValue(status);
  if (typeof notes === 'string') sh.getRange(row, 12).setValue(cleanText_(notes, 300));
}

/* ================= photos ================= */
function folder_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('PHOTO_FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder('Duni – תמונות לאתר');
  p.setProperty('PHOTO_FOLDER', f.getId());
  return f;
}
function upload_(d) {
  const m = /^data:(image\/(png|jpeg|webp));base64,(.+)$/.exec(String(d.data || ''));
  if (!m) throw new Error('bad_image');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[3]), m[1], (cleanText_(d.name, 40) || 'photo') + '.' + (m[2] === 'jpeg' ? 'jpg' : m[2]));
  const f = folder_().createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return f.getId();
}
function trashPhoto_(p) {
  if (!/^[A-Za-z0-9_-]{20,}$/.test(p)) return; // only photos uploaded from the admin page
  try { DriveApp.getFileById(p).setTrashed(true); } catch (e) {}
}

/* ================= helpers ================= */
function code_() {
  const abc = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let s = ''; for (let i = 0; i < 6; i++) s += abc[Math.floor(Math.random() * abc.length)];
  return s;
}
function stamp_(d) { return Utilities.formatDate(d || new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
function clean_(v, max) { return String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').replace(/^[=+\-@]/, "'$&").trim().slice(0, max); }
function cleanText_(v, max) { return String(v == null ? '' : v).replace(/\r/g, '').replace(/^[=+@]/, "'$&").trim().slice(0, max); }
function text_(s) { return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON); }
function json_(o) { return text_(JSON.stringify(o)); }
