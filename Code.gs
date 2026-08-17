/**
 * לוח ניהול החתונה – עומר & נועה
 * שרת Google Apps Script: מחבר בין הדשבורד (HTML) לגיליון בגוגל דרייב.
 *
 * התקנה (פעם אחת):
 *  1. פתח את תיקיית הדרייב שיצרת → חדש → Google Sheets (גיליון ריק).
 *  2. בגיליון: Extensions / תוספים ← Apps Script.
 *  3. מחק את מה שיש שם, הדבק את כל הקובץ הזה, ושמור.
 *  4. שנה את SECRET למטה לסיסמה משלך.
 *  5. הרץ את הפונקציה setup  (בוחרים אותה למעלה ולוחצים Run). אשר הרשאות.
 *  6. Deploy ← New deployment ← Web app
 *        Execute as:      Me
 *        Who has access:  Anyone
 *     העתק את כתובת ה-Web app.
 *  7. פתח את wedding-dashboard.html, לחץ "סנכרון וייבוא", הדבק את הכתובת + הסיסמה, ולחץ התחבר.
 *
 * אחרי כל שינוי בקובץ הזה צריך: Deploy ← Manage deployments ← עריכה (עיפרון) ← Version: New ← Deploy.
 */

/** סיסמה משותפת. כל מי שיש לו את הדשבורד צריך אותה. שנה אותה! */
const SECRET = 'omer-noa-2026';

/** שם תת-התיקייה בדרייב שבה יישמרו PDF-ים והצעות מחיר. */
const FILES_FOLDER_NAME = 'מסמכים וחוזים';

/* ────────────────────────────────────────────────────────────────
   מבנה הגיליון – כל לשונית והעמודות שלה.
   אפשר לערוך את התאים ישירות באקסל/שיטס, רק אל תשנה את שורת הכותרות
   ואל תמחק את עמודת id.
   ──────────────────────────────────────────────────────────────── */
const SCHEMA = {
  guests: {
    sheet: 'מוזמנים',
    cols: ['id', 'שם מלא', 'צד', 'קטגוריה', 'טלפון', 'מבוגרים', 'ילדים', 'אישור הגעה', 'סבירות הגעה', 'הגיע', 'מתנה ₪', 'שולחן', 'הערות'],
    keys: ['id', 'name', 'side', 'category', 'phone', 'adults', 'kids', 'rsvp', 'chance', 'arrived', 'gift', 'table', 'notes'],
    types: ['s', 's', 's', 's', 's', 'n', 'n', 's', 's', 'b', 'n', 's', 's']
  },
  vendors: {
    sheet: 'ספקים',
    cols: ['id', 'סוג', 'שם הספק', 'איש קשר', 'מקדמה', 'סכום כולל', 'מועד תשלום', 'אמצעי תשלום', 'סטטוס', 'הערות'],
    keys: ['id', 'type', 'name', 'contact', 'advance', 'total', 'when', 'method', 'status', 'notes'],
    types: ['s', 's', 's', 's', 'n', 'n', 's', 's', 's', 's']
  },
  venues: {
    sheet: 'מקומות',
    cols: ['id', 'שם המקום', 'תאריך פנוי', 'יום', 'מחיר למנה', 'מינימום אורחים', 'תוספות', 'הערות', 'סטטוס'],
    keys: ['id', 'name', 'date', 'day', 'pricePer', 'minGuests', 'extras', 'notes', 'status'],
    types: ['s', 's', 's', 's', 'n', 'n', 's', 's', 's']
  },
  tasks: {
    sheet: 'משימות',
    cols: ['id', 'משימה', 'קטגוריה', 'תאריך יעד', 'בוצע'],
    keys: ['id', 'title', 'category', 'due', 'done'],
    types: ['s', 's', 's', 'd', 'b']
  },
  tables: {
    sheet: 'שולחנות',
    cols: ['id', 'שם', 'צורה', 'מקומות', 'X', 'Y'],
    keys: ['id', 'name', 'shape', 'seats', 'x', 'y'],
    types: ['s', 's', 's', 'n', 'n', 'n']
  },
  files: {
    sheet: 'קבצים',
    cols: ['id', 'שייך ל', 'מזהה', 'שם הקובץ', 'קישור', 'סוג'],
    keys: ['id', 'ownerType', 'ownerId', 'name', 'url', 'mime'],
    types: ['s', 's', 's', 's', 's', 's']
  }
};

const SETTINGS_SHEET = 'הגדרות';

/* ════════════════════════════ נקודות כניסה ════════════════════════════ */

function doGet(e) {
  return json(route((e && e.parameter) || {}));
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { body = (e && e.parameter) || {}; }
  return json(route(body));
}

function route(req) {
  try {
    if (req.action === 'ping') return { ok: true, version: 1 };
    if (String(req.token || '') !== String(SECRET)) return { ok: false, error: 'סיסמה שגויה' };

    switch (req.action) {
      case 'load':    return { ok: true, data: readAll() };
      case 'rev':     return { ok: true, rev: getRev() };
      case 'save':    return saveAll(req);
      case 'upload':  return uploadFile(req);
      case 'delfile': return deleteFile(req);
      default:        return { ok: false, error: 'פעולה לא מוכרת: ' + req.action };
    }
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ════════════════════════════ קריאה ════════════════════════════ */

function readAll() {
  var ss = SpreadsheetApp.getActive();
  var out = { settings: readSettings() };
  Object.keys(SCHEMA).forEach(function (name) {
    out[name] = readSheet(ss, SCHEMA[name]);
  });
  out.rev = Number(out.settings.rev || 0);
  try {
    var parent = DriveApp.getFileById(ss.getId()).getParents().next();
    out.folderUrl = parent.getUrl();
    out.sheetUrl = ss.getUrl();
  } catch (err) {}
  return out;
}

function readSheet(ss, def) {
  var sh = ss.getSheetByName(def.sheet);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];

  // קוראים לפי שם הכותרת ולא לפי מיקום — כך הוספת עמודה או שינוי סדר
  // בגיליון לא מזיזה נתונים לעמודה הלא נכונה.
  var width = Math.min(Math.max(def.cols.length, sh.getLastColumn()), sh.getMaxColumns());
  var head = sh.getRange(1, 1, 1, width).getValues()[0].map(function (h) { return String(h).trim(); });
  var idx = def.cols.map(function (c) { return head.indexOf(c); });
  if (!idx.some(function (p) { return p >= 0; })) idx = def.cols.map(function (c, i) { return i; }); // אין כותרות – לפי מיקום

  var values = sh.getRange(2, 1, last - 1, width).getValues();
  var rows = [];
  values.forEach(function (r) {
    var empty = r.every(function (c) { return c === '' || c === null; });
    if (empty) return;
    var o = {};
    def.keys.forEach(function (k, i) {
      var p = idx[i];
      o[k] = decodeCell(p >= 0 ? r[p] : '', def.types[i]);
    });
    if (!o.id) o.id = uid();
    rows.push(o);
  });
  return rows;
}

function decodeCell(v, type) {
  if (type === 'n') return v === '' || v === null ? 0 : Number(v) || 0;
  if (type === 'b') {
    if (v === true) return true;
    var s = String(v).trim().toLowerCase();
    return s === 'כן' || s === 'v' || s === 'true' || s === 'x' || s === '✓';
  }
  if (type === 'd') {
    if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    return String(v || '').trim();
  }
  return v === null || v === undefined ? '' : String(v).trim();
}

function readSettings() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SETTINGS_SHEET);
  var o = {};
  if (!sh || sh.getLastRow() < 2) return o;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
    if (r[0] === '') return;
    var v = r[1];
    if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    o[String(r[0]).trim()] = v;
  });
  return o;
}

/* ════════════════════════════ שמירה ════════════════════════════ */

function saveAll(req) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, error: 'הגיליון תפוס, נסה שוב' };
  try {
    var current = getRev();
    // הדשבורד שולח את המספר שהיה לו בזמן הטעינה. אם בינתיים מישהו ערך
    // את הגיליון ידנית – לא דורסים אלא מחזירים התנגשות.
    if (!req.force && req.rev !== undefined && Number(req.rev) !== current) {
      return { ok: false, conflict: true, rev: current, error: 'הגיליון עודכן מבחוץ' };
    }

    var ss = SpreadsheetApp.getActive();
    var data = req.data || {};
    Object.keys(SCHEMA).forEach(function (name) {
      if (data[name]) writeSheet(ss, SCHEMA[name], data[name]);
    });
    if (data.settings) writeSettings(data.settings);

    var next = current + 1;
    setSetting('rev', next);
    setSetting('עודכן לאחרונה', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'));
    SpreadsheetApp.flush();
    return { ok: true, rev: next };
  } finally {
    lock.releaseLock();
  }
}

function writeSheet(ss, def, rows) {
  var sh = ss.getSheetByName(def.sheet) || ss.insertSheet(def.sheet);
  var width = def.cols.length;
  if (sh.getMaxColumns() < width) sh.insertColumnsAfter(sh.getMaxColumns(), width - sh.getMaxColumns());

  sh.getRange(1, 1, 1, width).setValues([def.cols]);

  var out = rows.map(function (r) {
    return def.keys.map(function (k, i) { return encodeCell(r[k], def.types[i]); });
  });

  var lastRow = sh.getMaxRows();
  if (lastRow > out.length + 1) {
    sh.getRange(out.length + 2, 1, lastRow - out.length - 1, width).clearContent();
  }
  if (out.length) {
    if (sh.getMaxRows() < out.length + 1) sh.insertRowsAfter(sh.getMaxRows(), out.length + 1 - sh.getMaxRows());
    sh.getRange(2, 1, out.length, width).setValues(out);
  }
}

function encodeCell(v, type) {
  if (type === 'n') return v === '' || v === null || v === undefined ? 0 : Number(v) || 0;
  if (type === 'b') return v ? 'כן' : '';
  return v === null || v === undefined ? '' : v;
}

function writeSettings(obj) {
  Object.keys(obj).forEach(function (k) {
    if (k === 'rev') return;
    setSetting(k, obj[k]);
  });
}

function setSetting(key, value) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SETTINGS_SHEET);
  if (!sh) { sh = SpreadsheetApp.getActive().insertSheet(SETTINGS_SHEET); sh.getRange(1, 1, 1, 2).setValues([['מפתח', 'ערך']]); }
  var last = Math.max(sh.getLastRow(), 1);
  var keys = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues() : [];
  for (var i = 0; i < keys.length; i++) {
    if (String(keys[i][0]).trim() === key) { sh.getRange(i + 2, 2).setValue(value); return; }
  }
  sh.getRange(last + 1, 1, 1, 2).setValues([[key, value]]);
}

function getRev() {
  return Number(readSettings().rev || 0);
}

/** נדלק כשעורכים את הגיליון ביד – כדי שהדשבורד יידע שיש חדש. */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    if (e.range.getSheet().getName() === SETTINGS_SHEET) return;
    setSetting('rev', getRev() + 1);
  } catch (err) { /* לא מפריעים לעריכה */ }
}

/* ════════════════════════════ קבצים בדרייב ════════════════════════════ */

function filesFolder() {
  var parent = DriveApp.getFileById(SpreadsheetApp.getActive().getId()).getParents().next();
  var it = parent.getFoldersByName(FILES_FOLDER_NAME);
  return it.hasNext() ? it.next() : parent.createFolder(FILES_FOLDER_NAME);
}

function uploadFile(req) {
  var folder = filesFolder();
  var blob = Utilities.newBlob(Utilities.base64Decode(req.data), req.mime || 'application/octet-stream', req.name || 'file');
  var f = folder.createFile(blob);
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) {}
  var row = {
    id: f.getId(),
    ownerType: req.ownerType || 'general',
    ownerId: req.ownerId || '',
    name: f.getName(),
    url: 'https://drive.google.com/file/d/' + f.getId() + '/view',
    mime: f.getMimeType()
  };
  appendRow(SCHEMA.files, row);
  setSetting('rev', getRev() + 1);
  return { ok: true, file: row, rev: getRev() };
}

function deleteFile(req) {
  try { DriveApp.getFileById(req.id).setTrashed(true); } catch (err) {}
  var ss = SpreadsheetApp.getActive();
  var rows = readSheet(ss, SCHEMA.files).filter(function (r) { return r.id !== req.id; });
  writeSheet(ss, SCHEMA.files, rows);
  setSetting('rev', getRev() + 1);
  return { ok: true, rev: getRev() };
}

function appendRow(def, obj) {
  var sh = SpreadsheetApp.getActive().getSheetByName(def.sheet);
  sh.appendRow(def.keys.map(function (k, i) { return encodeCell(obj[k], def.types[i]); }));
}

/* ════════════════════════════ בנייה ראשונית ════════════════════════════ */

function uid() { return Math.random().toString(36).slice(2, 9); }

/**
 * משלים עמודות שנוספו לסכמה אחרי שהגיליון כבר נוצר.
 * מוסיף עמודה ריקה במיקום הנכון ודוחף את שאר הנתונים ימינה,
 * כדי ששורות קיימות לא יזוזו לעמודה הלא נכונה.
 */
function alignColumns(sh, def) {
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) return;
  var head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  if (!head.some(function (h) { return h; })) return;   // אין שורת כותרות – אין מה ליישר
  def.cols.forEach(function (c, i) {
    if (head.indexOf(c) === -1) {
      sh.insertColumnBefore(i + 1);
      sh.getRange(1, i + 1).setValue(c);
      head.splice(i, 0, c);
    }
  });
}

/**
 * מריצים פעם אחת: בונה את כל הלשוניות, הכותרות, עיצוב ונתוני התחלה.
 * הרצה חוזרת לא מוחקת נתונים קיימים – רק משלימה לשוניות חסרות.
 */
function setup() {
  var ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetLocale('he_IL');

  Object.keys(SCHEMA).forEach(function (name) {
    var def = SCHEMA[name];
    var sh = ss.getSheetByName(def.sheet);
    var fresh = !sh;
    if (fresh) sh = ss.insertSheet(def.sheet);
    if (!fresh) alignColumns(sh, def);   // גיליון ותיק – משלימים עמודות חדשות במקום הנכון
    sh.setRightToLeft(true);
    sh.getRange(1, 1, 1, def.cols.length).setValues([def.cols])
      .setFontWeight('bold').setBackground('#3E5D50').setFontColor('#F5F2ED');
    sh.setFrozenRows(1);
    sh.hideColumns(1); // עמודת id – טכנית
    // עמודות טקסט (טלפון, איש קשר, id…) מקבלות פורמט "טקסט רגיל",
    // אחרת גוגל מזהה 0527283843 כמספר ובולעת את ה-0 המוביל.
    var rows = Math.max(sh.getMaxRows() - 1, 1);
    def.types.forEach(function (t, i) {
      if (t === 's') sh.getRange(2, i + 1, rows, 1).setNumberFormat('@');
    });
    if (fresh && SEED[name]) SEED[name].forEach(function (r) { appendRow(def, r); });
    sh.autoResizeColumns(2, def.cols.length - 1);
  });

  var st = ss.getSheetByName(SETTINGS_SHEET);
  var freshSettings = !st;
  if (freshSettings) st = ss.insertSheet(SETTINGS_SHEET);
  st.setRightToLeft(true);
  st.getRange(1, 1, 1, 2).setValues([['מפתח', 'ערך']])
    .setFontWeight('bold').setBackground('#3E5D50').setFontColor('#F5F2ED');
  st.setFrozenRows(1);
  st.setColumnWidth(1, 160); st.setColumnWidth(2, 200);
  // משלימים רק ערכים שחסרים — לא דורסים מה שכבר הוגדר
  var have = readSettings();
  if (!have['תאריך החתונה']) setSetting('תאריך החתונה', '2026-12-17');
  if (!have['תקציב מתוכנן']) setSetting('תקציב מתוכנן', 160000);
  if (!have.rev) setSetting('rev', 1);
  setSetting('עודכן לאחרונה', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'));

  var def0 = ss.getSheetByName('Sheet1') || ss.getSheetByName('גיליון1');
  if (def0 && ss.getSheets().length > 1) ss.deleteSheet(def0);

  ss.setActiveSheet(ss.getSheetByName(SCHEMA.guests.sheet));
  filesFolder();
  SpreadsheetApp.getUi().alert('הגיליון מוכן ✓\n\nעכשיו: Deploy ← New deployment ← Web app\nExecute as: Me | Who has access: Anyone');
}

/** נתוני התחלה – אותם נתונים שהיו בדשבורד. אפשר למחוק הכל ולהזין את שלכם. */
var SEED = {
  guests: [
    { id: uid(), name: 'משפחת לוי - דוד ורונית', side: 'עומר', category: 'משפחה', phone: '052-4412093', adults: 2, kids: 0, rsvp: 'אישר', arrived: false, gift: 0, table: '', notes: 'הורים של עומר' },
    { id: uid(), name: 'יעל לוי', side: 'עומר', category: 'משפחה', phone: '054-7781230', adults: 1, kids: 0, rsvp: 'אישר', arrived: false, gift: 0, table: '', notes: 'אחות' },
    { id: uid(), name: 'סבתא מרים', side: 'עומר', category: 'משפחה', phone: '03-6721145', adults: 1, kids: 0, rsvp: 'אישר', arrived: false, gift: 0, table: '', notes: 'להושיב קרוב לבמה' },
    { id: uid(), name: 'משפחת ברקוביץ', side: 'עומר', category: 'משפחה מורחבת', phone: '050-3312876', adults: 2, kids: 2, rsvp: 'טרם', arrived: false, gift: 0, table: '', notes: '' },
    { id: uid(), name: 'אורי כהן', side: 'עומר', category: 'צבא', phone: '052-9983311', adults: 2, kids: 0, rsvp: 'אישר', arrived: false, gift: 0, table: '', notes: '' },
    { id: uid(), name: 'רון פרץ', side: 'עומר', category: 'צבא', phone: '053-2214467', adults: 1, kids: 0, rsvp: 'טרם', arrived: false, gift: 0, table: '', notes: '' },
    { id: uid(), name: 'נדב שגיא', side: 'עומר', category: 'עבודה', phone: '054-1120934', adults: 2, kids: 0, rsvp: 'אישר', arrived: false, gift: 0, table: '', notes: 'מנהל הצוות' },
    { id: uid(), name: 'שירה אלון', side: 'עומר', category: 'עבודה', phone: '058-7712340', adults: 1, kids: 0, rsvp: 'טרם', arrived: false, gift: 0, table: '', notes: '' },
    { id: uid(), name: 'גיא ומאיה', side: 'עומר', category: 'חברים', phone: '050-4483321', adults: 2, kids: 1, rsvp: 'אישר', arrived: false, gift: 0, table: '', notes: '' },
    { id: uid(), name: 'משפחת אזולאי', side: 'נועה', category: 'משפחה', phone: '052-3341290', adults: 2, kids: 0, rsvp: 'אישר', arrived: false, gift: 0, table: '', notes: 'הורים של נועה' },
    { id: uid(), name: 'תמר אזולאי', side: 'נועה', category: 'משפחה', phone: '054-9987120', adults: 1, kids: 0, rsvp: 'אישר', arrived: false, gift: 0, table: '', notes: 'אחות + עדה' },
    { id: uid(), name: 'דודה אילנה', side: 'נועה', category: 'משפחה מורחבת', phone: '09-7745120', adults: 2, kids: 0, rsvp: 'טרם', arrived: false, gift: 0, table: '', notes: '' },
    { id: uid(), name: 'נועם ברק', side: 'נועה', category: 'ילדות', phone: '050-8871233', adults: 2, kids: 0, rsvp: 'אישר', arrived: false, gift: 0, table: '', notes: '' },
    { id: uid(), name: 'הילה מזרחי', side: 'נועה', category: 'ילדות', phone: '053-4419087', adults: 1, kids: 0, rsvp: 'אישר', arrived: false, gift: 0, table: '', notes: '' },
    { id: uid(), name: 'ליאור ואפרת', side: 'נועה', category: 'חברים', phone: '052-6612088', adults: 2, kids: 2, rsvp: 'טרם', arrived: false, gift: 0, table: '', notes: '' },
    { id: uid(), name: 'דנה שטרן', side: 'נועה', category: 'עבודה', phone: '054-3390127', adults: 1, kids: 0, rsvp: 'לא מגיע', arrived: false, gift: 0, table: '', notes: 'בחו״ל' },
    { id: uid(), name: 'מיכל רוזן', side: 'נועה', category: 'עבודה', phone: '058-2213099', adults: 2, kids: 0, rsvp: 'אישר', arrived: false, gift: 0, table: '', notes: '' },
    { id: uid(), name: 'משפחת שוורץ', side: 'נועה', category: 'שכונה', phone: '050-1123998', adults: 2, kids: 3, rsvp: 'טרם', arrived: false, gift: 0, table: '', notes: '' }
  ],
  tables: [
    { id: uid(), name: 'שולחן 1', shape: 'עגול', seats: 12, x: 60, y: 50 },
    { id: uid(), name: 'שולחן 2', shape: 'עגול', seats: 12, x: 280, y: 50 },
    { id: uid(), name: 'שולחן 3', shape: 'מרובע', seats: 8, x: 500, y: 60 },
    { id: uid(), name: 'שולחן אבירים', shape: 'ארוך', seats: 20, x: 60, y: 300 }
  ],
  vendors: [
    { id: uid(), type: 'די ג׳יי', name: 'פרץ גלאס', contact: '052-5446859', advance: 1000, total: 8500, when: 'עד שבוע אחרי', method: 'העברה בנקאית', status: 'נסגר', notes: '' },
    { id: uid(), type: 'צלם סטילס', name: 'עידן מרציאנו', contact: '054-2523442', advance: 1500, total: 7300, when: 'חצי ביומיים אחרי', method: 'העברה', status: 'נסגר', notes: 'כולל אלבום' },
    { id: uid(), type: 'איפור ושיער', name: 'נוי שילר', contact: '054-8195800', advance: 300, total: 3400, when: 'ביום האירוע', method: 'ביט', status: 'נסגר', notes: 'כולל ליווי' },
    { id: uid(), type: 'עיצוב ופרחים', name: 'סטודיו ורד', contact: '050-7781203', advance: 0, total: 6200, when: 'שבוע לפני', method: 'העברה', status: 'בבירור', notes: '' },
    { id: uid(), type: 'רב', name: 'הרב שמעוני', contact: '052-3312099', advance: 0, total: 1500, when: 'ביום האירוע', method: 'מזומן', status: 'בבירור', notes: '' }
  ],
  venues: [
    { id: uid(), name: 'בית רבן', date: '3.06', day: 'רביעי', pricePer: 520, minGuests: 300, extras: '13,500 תאורה והגברה', notes: 'בר בתוספת 39 לאורח', status: 'ביקרנו' },
    { id: uid(), name: 'בית ברל', date: '28.06', day: 'ראשון', pricePer: 450, minGuests: 280, extras: '', notes: 'גן פתוח', status: 'מועדף' },
    { id: uid(), name: 'מחסיה', date: '19.06', day: 'שישי', pricePer: 500, minGuests: 300, extras: '', notes: 'רחוק יחסית', status: 'לבדיקה' },
    { id: uid(), name: 'חדרה - צהריים', date: 'ספטמבר', day: 'שישי', pricePer: 380, minGuests: 250, extras: '', notes: 'שישי צהריים בלבד', status: 'לבדיקה' }
  ],
  tasks: [
    { id: uid(), title: 'לסגור אולם ולחתום חוזה', category: 'אולם', due: '2026-09-15', done: false },
    { id: uid(), title: 'לקבוע טעימות', category: 'אולם', due: '2026-09-28', done: false },
    { id: uid(), title: 'לסיים רשימת מוזמנים סופית', category: 'אורחים', due: '2026-10-10', done: false },
    { id: uid(), title: 'להזמין הזמנות דיגיטליות', category: 'אורחים', due: '2026-10-20', done: false },
    { id: uid(), title: 'מדידות שמלה', category: 'לבוש', due: '2026-09-05', done: true },
    { id: uid(), title: 'חליפה לחתן', category: 'לבוש', due: '2026-10-01', done: false },
    { id: uid(), title: 'לסגור צלם וידאו', category: 'ספקים', due: '2026-09-20', done: false },
    { id: uid(), title: 'לתאם עם הרב', category: 'טקס', due: '2026-11-01', done: false },
    { id: uid(), title: 'רישום רבנות', category: 'טקס', due: '2026-08-30', done: false }
  ],
  files: []
};

// סבירות הגעה התחלתית נגזרת מאישור ההגעה; אפשר לשנות ידנית בכל שורה.
SEED.guests.forEach(function (g) {
  g.chance = g.rsvp === 'אישר' ? 'גבוהה' : g.rsvp === 'לא מגיע' ? 'נמוכה' : 'בינונית';
});
