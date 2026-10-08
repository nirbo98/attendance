/**
 * מערכת נוכחות — Google Apps Script
 *
 * שכבות ההגנה:
 * 1. קוד QR על המקרן שמתחלף כל 10 שניות וחתום בסוד שנמצא רק בשרת.
 *    צילום מסך שנשלח לחבר פג תוקף תוך שניות.
 * 2. רשימת סטודנטים סגורה: אפשר להירשם רק עם ת"ז שנמצאת ברשימת הקורס.
 * 3. קשירת מכשיר: כל טלפון קשור לסטודנט אחד. החלפת מכשיר אפשרית, אבל מסומנת לבדיקה.
 * 4. מיקום כסימון בלבד: המערכת לומדת לבד איפה הכיתה (החציון של הרישומים בסבב)
 *    ומסמנת רק מי שרחוק בוודאות. מי שלא שיתף מיקום פשוט לא מסומן.
 * 5. סבב בדיקה בזמן אקראי, ובדיקת פתע של שמות מוקרנים (מסומנים קודם).
 *    מי שלא נמצא בכיתה מאבד את הנוכחות של אותו שיעור בלבד.
 */

// ===== הגדרות =====
const CFG = {
  ROTATE_SECONDS: 10,          // כל כמה שניות הקוד על המסך מתחלף
  ACCEPT_PREVIOUS: 1,          // כמה קודים קודמים עדיין מתקבלים (1 = חלון של 10 עד 20 שניות)
  PASS_TTL_SECONDS: 180,       // כמה זמן יש לסטודנט להשלים רישום אחרי סריקה תקינה
  ROUND_MINUTES: 3,            // סבב נסגר לבד אחרי כמה דקות
  SPOT_CHECK_SIZE: 10,         // כמה שמות בכל בדיקת פתע
  ATTENDANCE_THRESHOLD: 0.8,   // מתחת לסף הזה הסטודנט מסומן באדום בדוח
  FAR_METERS: 1000,            // מרחק מהכיתה שממנו רישום מסומן כחשוד
  GEO_GOOD_ACCURACY: 250,      // רק מדידות מדויקות מזה משמשות לחישוב מיקום הכיתה
  GEO_MIN_POINTS: 8,           // מינימום מדידות טובות בסבב כדי לחשב מיקום כיתה
  CACHE_TTL: 21600,            // 6 שעות
  MIN_PIN_LENGTH: 8,
  TZ: 'Asia/Jerusalem',
  // דף הסטודנטים החיצוני (GitHub Pages). אין בו נתונים או סודות, הוא רק שולח בקשות לשרת הזה.
  // מחרוזת ריקה = משתמשים רק בדף הגיבוי של Apps Script.
  FRONTEND_URL: 'https://nirbo98.github.io/attendance/'
};

const SH = {
  START: 'מסך המרצה',
  STUDENTS: 'סטודנטים',
  SESSIONS: 'שיעורים',
  CHECKINS: 'רישומים',
  SPOT: 'בדיקות פתע',
  REPORT: 'דוח נוכחות'
};

const HEADERS = {
  'סטודנטים': ['ת"ז', 'שם מלא', 'מזהה מכשיר', 'נקשר בתאריך', 'החלפות מכשיר'],
  'שיעורים': ['מזהה שיעור', 'תאריך', 'כותרת', 'מספר סבבים', 'נוצר'],
  'רישומים': ['זמן', 'שיעור', 'סבב', 'ת"ז', 'שם', 'מכשיר', 'סימון חשד', 'סטטוס', 'קו רוחב', 'קו אורך', 'דיוק (מטר)'],
  'בדיקות פתע': ['זמן', 'שיעור', 'ת"ז', 'שם', 'נמצא בכיתה', 'היה מסומן']
};

const STATUS_OK = 'תקין';
const STATUS_VOID = 'נפסל – בדיקת פתע';
const FLAG_REBIND = 'החלפת מכשיר';
const FLAG_FAR = 'רחוק מהכיתה';

// ===== תפריט בגיליון =====
function onOpen() {
  SpreadsheetApp.getUi().createMenu('נוכחות')
    .addItem('הגדרה ראשונית', 'setupMenu')
    .addItem('הקישור למסך המרצה', 'linksMenu')
    .addItem('עדכון כתובת המערכת', 'setUrlMenu')
    .addSeparator()
    .addItem('שינוי קוד מנהל', 'changePinMenu')
    .addItem('הפקת דוח נוכחות', 'reportMenu')
    .addSeparator()
    .addItem('דף גיבוי (חירום בלבד)', 'backupMenu')
    .addToUi();
}

function setupMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // עותק חדש של הגיליון: מתחילים מאפס, כדי שסוד החתימה וקוד המנהל לא יועברו מהמקור
  if (props_().getProperty('SHEET_ID') !== ss.getId()) {
    ['SECRET', 'ADMIN_PIN_HASH', 'ACTIVE', 'WEBAPP_URL', 'FRONTEND'].forEach(k => props_().deleteProperty(k));
  }
  setup_(ss.getId());
  if (!props_().getProperty('ADMIN_PIN_HASH')) changePinMenu();
  SpreadsheetApp.getUi().alert(
    'ההגדרה הושלמה.\n\n' +
    'השלב הבא: הפעלת המערכת (פריסה), לפי המדריך. אחרי הפריסה בוחרים כאן בתפריט "הקישור למסך המרצה" ומדביקים את הכתובת שהתקבלה.'
  );
}

// גוגל לא תמיד מחזיר את הכתובת הציבורית של המערכת (לפעמים מחזיר כתובת פיתוח שנפתחת רק לבעלים).
// לכן הכתובת נשמרת פעם אחת, בהדבקה מחלון הפריסה.
const WEBAPP_RX = /^https:\/\/script\.google\.com\/(?:a\/macros\/([^/]+)|macros)\/s\/([\w-]+)\/exec$/;

function setUrlMenu() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('כתובת המערכת',
    'הדביקי כאן את הכתובת שהופיעה בסוף הפריסה (אפליקציית אינטרנט). היא מתחילה ב-https://script.google.com ומסתיימת ב-/exec',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return false;
  const url = String(res.getResponseText() || '').trim().replace(/[?#].*$/, '');
  if (!WEBAPP_RX.test(url)) {
    ui.alert('זו לא הכתובת הנכונה. צריך את הכתובת שמסתיימת ב-/exec, מחלון "פריסה" ב-Apps Script.');
    return false;
  }
  props_().setProperty('WEBAPP_URL', url);
  try { writeStartTab_(SpreadsheetApp.getActiveSpreadsheet()); } catch (e) { console.error(e); }
  return true;
}

function linksMenu() {
  const ui = SpreadsheetApp.getUi();
  if (!props_().getProperty('WEBAPP_URL') && !setUrlMenu()) return;
  const admin = props_().getProperty('WEBAPP_URL') + '?admin';
  const html = HtmlService.createHtmlOutput(
    '<div dir="rtl" style="font-family:Arial,sans-serif;font-size:14px;line-height:1.6">' +
    '<p>זה הקישור למסך המרצה. שמרי אותו כסימנייה במחשב שמחובר למקרן, ואפשר גם בטלפון:</p>' +
    '<p><a href="' + admin + '" target="_blank" style="word-break:break-all">' + admin + '</a></p>' +
    '<p style="color:#666">אם נפתחת שגיאה "לא ניתן לפתוח את הקובץ", זה בגלל כמה חשבונות גוגל מחוברים בדפדפן. פתחי את הקישור בחלון גלישה בסתר.</p>' +
    '<p style="color:#666">לסטודנטים לא צריך לשלוח קישור. הקוד שעל המסך מוביל אותם לדף הנכון.</p></div>'
  ).setWidth(460).setHeight(290);
  ui.showModalDialog(html, 'הקישור למסך המרצה');
}

// מעבר בין דף הסטודנטים הרגיל (GitHub) לדף הגיבוי של Apps Script.
// רק מהגיליון ולא ממסך המרצה, כדי שאף אחד לא יפעיל אותו בטעות באמצע שיעור.
function backupMenu() {
  const ui = SpreadsheetApp.getUi();
  if (frontendOn_()) {
    const r = ui.alert('דף גיבוי',
      'להעביר את הסטודנטים לדף הגיבוי?\n\nרק אם הדף הרגיל לא נטען לסטודנטים. בדף הגיבוי, מי שמחובר בטלפון לכמה חשבונות גוגל לא יוכל להירשם.',
      ui.ButtonSet.YES_NO);
    if (r === ui.Button.YES) { props_().setProperty('FRONTEND', 'off'); ui.alert('הסטודנטים מופנים עכשיו לדף הגיבוי.'); }
  } else {
    const r = ui.alert('דף גיבוי', 'המערכת נמצאת עכשיו בדף הגיבוי. להחזיר לדף הרגיל?', ui.ButtonSet.YES_NO);
    if (r === ui.Button.YES) { props_().setProperty('FRONTEND', 'on'); ui.alert('הסטודנטים מופנים שוב לדף הרגיל.'); }
  }
}

function changePinMenu() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('קוד מנהל', 'בחרי קוד כניסה למסך המרצה (לפחות ' + CFG.MIN_PIN_LENGTH + ' תווים):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  try {
    setPin_(res.getResponseText());
    ui.alert('הקוד נשמר.');
  } catch (e) {
    ui.alert('הקוד חייב להכיל לפחות ' + CFG.MIN_PIN_LENGTH + ' תווים.');
  }
}

function reportMenu() {
  const n = buildReport_();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.REPORT).activate();
  SpreadsheetApp.getUi().alert('הדוח עודכן (' + n + ' שיעורים).');
}

function setup_(sheetId) {
  const P = props_();
  P.setProperty('SHEET_ID', sheetId);
  if (!P.getProperty('SECRET')) P.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  const ss = SpreadsheetApp.openById(sheetId);
  Object.keys(HEADERS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
  });
  writeStartTab_(ss);
  // ת"ז נשמרת כטקסט כדי לא לאבד אפסים מובילים
  ss.getSheetByName(SH.STUDENTS).getRange('A:A').setNumberFormat('@');
  ss.getSheetByName(SH.CHECKINS).getRange('D:D').setNumberFormat('@');
  ss.getSheetByName(SH.SPOT).getRange('C:C').setNumberFormat('@');
}

// לשונית ראשונה בגיליון עם קישור גדול למסך המרצה, כדי שלא יהיה צורך בסימנייה בדפדפן
function writeStartTab_(ss) {
  let sh = ss.getSheetByName(SH.START);
  if (!sh) sh = ss.insertSheet(SH.START, 0);
  try { sh.setRightToLeft(true); } catch (e) {}
  sh.clear();
  const url = props_().getProperty('WEBAPP_URL');
  sh.getRange('A1').setValue('מערכת נוכחות');
  if (url) {
    sh.getRange('A3').setFormula('=HYPERLINK("' + url + '?admin", "לחצי כאן לפתיחת מסך המרצה")');
    sh.getRange('A5').setValue('אם נפתח דף של גוגל עם "לא ניתן לפתוח את הקובץ": פתחי חלון גלישה בסתר והדביקי בו את הכתובת הזו:');
    sh.getRange('A6').setValue(url + '?admin');
  } else {
    sh.getRange('A3').setValue('הקישור יופיע כאן אחרי ההגדרה (תפריט נוכחות, "עדכון כתובת המערכת").');
  }
  sh.getRange('A8').setValue('הדוח נמצא בלשונית "דוח נוכחות". את שאר הלשוניות עדיף לא לשנות.');
  try {
    sh.getRange('A1').setFontSize(20).setFontWeight('bold');
    sh.getRange('A3').setFontSize(18).setFontWeight('bold');
    sh.setColumnWidth(1, 700);
  } catch (e) {}
}

function setPin_(pin) {
  pin = String(pin || '').trim();
  if (pin.length < CFG.MIN_PIN_LENGTH) throw E_('shortpin');
  props_().setProperty('ADMIN_PIN_HASH', pinHash_(pin));
}

// ===== כניסה לאתר =====
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.admin !== undefined) return page_('Admin', {}, 'ניהול נוכחות');
  let boot;
  if (p.s && p.r && p.c) boot = validateScan_(String(p.s), Number(p.r), String(p.c));
  else boot = { state: 'noscan' };
  return page_('Student', boot, 'רישום נוכחות');
}

function page_(file, boot, title) {
  const t = HtmlService.createTemplateFromFile(file);
  t.boot = JSON.stringify(boot).replace(/</g, '\\u003c');
  return t.evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

// ===== ממשק לדף הסטודנטים החיצוני =====
// רק פעולות סטודנט. פעולות מרצה זמינות רק ממסך המרצה, עם קוד מנהל.
const STUDENT_API = {
  scan: req => validateScan_(String(req.s || ''), Number(req.r), String(req.c || '')),
  checkIn: checkIn,
  lookupId: lookupId,
  register: register,
  attachLocation: attachLocation,
  myAttendance: myAttendance
};

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = Object.prototype.hasOwnProperty.call(STUDENT_API, req.action) ? STUDENT_API[req.action] : null;
    out = fn ? fn(req) : { error: 'badaction' };
  } catch (err) {
    out = { error: 'server' };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function validateScan_(s, r, c) {
  const a = active_();
  if (!a || !a.open) return { state: 'closed' };
  if (a.s !== s || a.r !== r) return { state: 'expired' };
  const now = slot_();
  let ok = false;
  for (let k = 0; k <= CFG.ACCEPT_PREVIOUS; k++) {
    if (safeEq_(code_(s, r, now - k), c)) ok = true;
  }
  if (!ok) return { state: 'expired' };
  const issued = Date.now();
  return { state: 'ok', s: s, r: r, issued: issued, pass: passFor_(s, r, issued), label: a.label || '' };
}

// ===== פעולות הסטודנט =====
function checkIn(req) {
  return api_(() => {
    const auth = verifyPass_(req);
    const token = cleanToken_(req.token);
    if (!token) return { needRegister: true };
    const st = loadStudents_().find(x => x.token === token);
    if (!st) return { needRegister: true };
    return withLock_(() => record_(st, auth.s, auth.r, []));
  });
}

function lookupId(req) {
  return api_(() => {
    verifyPass_(req);
    const id = normId_(req.id);
    if (!id) throw E_('badid');
    const st = loadStudents_().find(x => x.id === id);
    if (!st) throw E_('notfound');
    return { name: st.name };
  });
}

function register(req) {
  return api_(() => {
    const auth = verifyPass_(req);
    const token = cleanToken_(req.token);
    if (!token) throw E_('badtoken');
    const id = normId_(req.id);
    if (!id) throw E_('badid');
    return withLock_(() => {
      const all = loadStudents_();
      const st = all.find(x => x.id === id);
      if (!st) throw E_('notfound');
      const owner = all.find(x => x.token === token);
      if (owner && owner.id !== st.id) throw E_('devicetaken', owner.name);
      if (st.token === token) return record_(st, auth.s, auth.r, []);
      // מישהו כבר רשם את הת"ז הזו בסבב הנוכחי ממכשיר אחר: לא מאפשרים "לגנוב" את הרישום
      if (alreadyIn_(auth.s, auth.r, st.id)) throw E_('alreadyother');
      const flags = [];
      if (st.token) flags.push(FLAG_REBIND);
      const rebinds = st.token ? st.rebinds + 1 : st.rebinds;
      sheet_(SH.STUDENTS).getRange(st.row, 3, 1, 3).setValues([[token, new Date(), rebinds]]);
      st.token = token;
      return record_(st, auth.s, auth.r, flags);
    });
  });
}

// המיקום נשלח בנפרד אחרי האישור, כדי שהרישום עצמו לא יחכה לו
function attachLocation(req) {
  return api_(() => {
    const auth = verifyPass_(req);
    const a = active_();
    if (!a || a.s !== auth.s) return { ok: true };   // שיעור שכבר הסתיים: לא שומרים מיקום
    const token = cleanToken_(req.token);
    const st = token && loadStudents_().find(x => x.token === token);
    if (!st) throw E_('notfound');
    const row = Math.floor(Number(req.row));
    const lat = Number(req.lat), lng = Number(req.lng), acc = Number(req.acc);
    if (!(row >= 2) || !isFinite(lat) || !isFinite(lng) || !isFinite(acc) ||
        Math.abs(lat) > 90 || Math.abs(lng) > 180 || acc <= 0 || acc > 100000) throw E_('badloc');
    const sh = sheet_(SH.CHECKINS);
    if (row > sh.getLastRow()) throw E_('badloc');
    const v = sh.getRange(row, 2, 1, 10).getValues()[0];
    if (String(v[0]) !== auth.s || Number(v[1]) !== auth.r || normId_(v[2]) !== st.id) throw E_('badloc');
    if (v[7] !== '' && v[7] !== null) return { ok: true };
    // עיגול לכ-10 מטרים: מספיק לבדיקה, לא יותר מדויק מהנדרש
    sh.getRange(row, 9, 1, 3).setValues([[Math.round(lat * 1e4) / 1e4, Math.round(lng * 1e4) / 1e4, Math.round(acc)]]);
    return { ok: true };
  });
}

function myAttendance(req) {
  return api_(() => {
    verifyPass_(req);
    const token = cleanToken_(req.token);
    const st = token && loadStudents_().find(x => x.token === token);
    if (!st) throw E_('notfound');
    const d = attendanceData_();
    let full = 0;
    d.sessions.forEach(se => { if (cellFor_(d, se, st.id) === '✓') full++; });
    return { full: full, total: d.sessions.length };
  });
}

function record_(st, s, r, flags) {
  const c = cache_();
  const key = ['ck', s, r, st.id].join('|');
  if (c.get(key)) return { ok: true, already: true, name: st.name, r: r };
  if (alreadyInSheet_(s, r, st.id)) {
    c.put(key, '1', CFG.CACHE_TTL);
    return { ok: true, already: true, name: st.name, r: r };
  }
  // המונים נקראים לפני ההוספה, כדי שחישוב מהגיליון (כשהמטמון ריק) לא יספור את השורה החדשה פעמיים
  const cnt = count_(s, r);
  const fl = flags.length ? flagCount_(s) : 0;
  const sh = sheet_(SH.CHECKINS);
  sh.appendRow([new Date(), s, r, st.id, st.name, st.token.slice(0, 8), flags.join(', '), STATUS_OK, '', '', '']);
  const row = sh.getLastRow();
  c.put(key, '1', CFG.CACHE_TTL);
  c.put(['cnt', s, r].join('|'), String(cnt + 1), CFG.CACHE_TTL);
  if (flags.length) c.put('fl|' + s, String(fl + 1), CFG.CACHE_TTL);
  return { ok: true, name: st.name, r: r, row: row, flagged: flags.length > 0 };
}

// ===== פעולות המרצה =====
function adminLogin(pin) {
  return adminApi_(pin, () => {
    const st = loadStudents_();
    return {
      active: active_(),
      rosterSize: st.length,
      registered: st.filter(x => x.token).length,
      rotate: CFG.ROTATE_SECONDS,
      spotSize: CFG.SPOT_CHECK_SIZE,
      roundMinutes: CFG.ROUND_MINUTES,
      urlMissing: !baseUrl_(),
      frontendAvailable: !!CFG.FRONTEND_URL && !!apiKey_(),
      frontendOn: frontendOn_()
    };
  });
}

function startSession(pin, label) {
  const res = startSessionLocked_(pin, label);
  // הדוח מתעדכן לבד בתחילת כל שיעור, כך שהוא תמיד נכון עד השיעור הקודם.
  // תקלה בדוח לעולם לא עוצרת את פתיחת השיעור.
  if (res && !res.error) { try { buildReport_(); } catch (e) { console.error(e); } }
  return res;
}

function startSessionLocked_(pin, label) {
  return adminApi_(pin, () => withLock_(() => {
    const cur = active_();
    if (cur && cur.open) throw E_('roundopen');
    scrubLocations_();
    const sh = sheet_(SH.SESSIONS);
    const existing = new Set(colValues_(sh, 1).map(String));
    const base = Utilities.formatDate(new Date(), CFG.TZ, 'yyyyMMdd-HHmm');
    let s = base, n = 2;
    while (existing.has(s)) s = base + '_' + (n++);
    label = String(label || '').slice(0, 80);
    sh.appendRow([s, new Date(), label, 0, new Date()]);
    const a = { s: s, r: 0, open: false, label: label };
    setActive_(a);
    return a;
  }));
}

function openRound(pin) {
  return adminApi_(pin, () => withLock_(() => {
    const a = active_();
    if (!a) throw E_('nosession');
    if (a.open) return a;
    a.r += 1;
    a.open = true;
    a.openedAt = Date.now();
    setActive_(a);
    setSessionRounds_(a.s, a.r);
    return a;
  }));
}

function closeRound(pin) {
  return adminApi_(pin, () => {
    const a = active_();
    if (!a) throw E_('nosession');
    a.open = false;
    setActive_(a);
    return a;
  });
}

function endSession(pin) {
  return adminApi_(pin, () => {
    props_().deleteProperty('ACTIVE');
    return { ended: true };
  });
}

function getCode(pin) {
  return adminApi_(pin, () => {
    const a = active_();
    if (!a || !a.open) return { open: false };
    if (!baseUrl_()) throw E_('nourl');
    const ms = CFG.ROTATE_SECONDS * 1000;
    const sl = slot_();
    return {
      open: true,
      r: a.r,
      url: studentUrl_(a, sl),
      msLeft: (sl + 1) * ms - Date.now(),
      roundMsLeft: a.openedAt + CFG.ROUND_MINUTES * 60000 - Date.now(),
      count: count_(a.s, a.r)
    };
  });
}

function getStats(pin) {
  return adminApi_(pin, () => {
    const a = active_();
    if (!a) return { active: null };
    const rows = checkinRows_(a.s).filter(x => x.status === STATUS_OK);
    const rounds = [];
    for (let r = 1; r <= a.r; r++) rounds.push({ r: r, count: count_(a.s, r) });
    const far = farRows_(rows);
    const farIds = new Set(rows.filter(x => far.has(x.row)).map(x => x.id));
    const rebindIds = new Set(rows.filter(x => x.flags).map(x => x.id));
    return { active: a, rounds: rounds, rebind: rebindIds.size, far: farIds.size };
  });
}

function spotCheck(pin) {
  return adminApi_(pin, () => {
    const a = active_();
    if (!a) throw E_('nosession');
    const done = new Set(spotRows_(a.s).map(x => x.id));
    const rows = checkinRows_(a.s).filter(x => x.status === STATUS_OK);
    const far = farRows_(rows);
    const byId = {};
    rows.forEach(x => {
      const o = byId[x.id] || (byId[x.id] = { id: x.id, name: x.name, flagged: false });
      if (x.flags || far.has(x.row)) o.flagged = true;
    });
    const pool = Object.keys(byId).map(k => byId[k]).filter(x => !done.has(x.id));
    shuffle_(pool);
    pool.sort((x, y) => (y.flagged ? 1 : 0) - (x.flagged ? 1 : 0));
    return pool.slice(0, CFG.SPOT_CHECK_SIZE);
  });
}

function markSpot(pin, id, present) {
  return adminApi_(pin, () => withLock_(() => {
    const a = active_();
    if (!a) throw E_('nosession');
    id = normId_(id);
    const all = checkinRows_(a.s);
    const rows = all.filter(x => x.id === id);
    if (!rows.length) throw E_('notfound');
    const far = farRows_(all.filter(x => x.status === STATUS_OK));
    const flagged = rows.some(x => x.flags || far.has(x.row));
    sheet_(SH.SPOT).appendRow([new Date(), a.s, id, rows[0].name, present ? 'כן' : 'לא', flagged ? 'כן' : 'לא']);
    if (!present) {
      const sh = sheet_(SH.CHECKINS);
      rows.forEach(x => sh.getRange(x.row, 8).setValue(STATUS_VOID));
    }
    return { id: id, present: !!present };
  }));
}

function resetDevice(pin, id) {
  return adminApi_(pin, () => withLock_(() => {
    id = normId_(id);
    const st = loadStudents_().find(x => x.id === id);
    if (!st) throw E_('notfound');
    sheet_(SH.STUDENTS).getRange(st.row, 3).setValue('');
    return { name: st.name };
  }));
}

function makeReport(pin) {
  return adminApi_(pin, () => ({ sessions: buildReport_() }));
}

// ===== ייבוא רשימה מהמודל (הדבקה מאקסל או CSV) =====
function importRoster(pin, text) {
  return adminApi_(pin, () => withLock_(() => {
    const parsed = parseRoster_(String(text || ''));
    const existing = {};
    loadStudents_().forEach(s => { existing[s.id] = s; });
    const sh = sheet_(SH.STUDENTS);
    const seen = {};
    const add = [];
    let updated = 0;
    parsed.list.forEach(p => {
      const id = normId_(p.rawId);
      if (seen[id]) return;
      seen[id] = true;
      const ex = existing[id];
      if (ex) {
        if (p.name && p.name !== ex.name) { sh.getRange(ex.row, 2).setValue(p.name); updated++; }
      } else {
        add.push([p.rawId, p.name]);
      }
    });
    if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 2).setValues(add);
    return { added: add.length, updated: updated, skipped: parsed.skipped, total: Object.keys(existing).length + add.length };
  }));
}

function parseRoster_(text) {
  const lines = text.replace(/\r/g, '').split('\n').filter(l => l.trim());
  if (!lines.length) throw E_('empty');
  const delim = lines.some(l => l.indexOf('\t') >= 0) ? '\t' : (lines[0].indexOf(';') >= 0 && lines[0].indexOf(',') < 0 ? ';' : ',');
  const rows = lines.map(l => splitLine_(l, delim).map(c => c.trim()));
  const ncol = Math.max.apply(null, rows.map(r => r.length));
  if (ncol < 2) return parseLoose_(lines, RX_ROSTER_());

  const RX = RX_ROSTER_();
  const looksId = v => { const d = normId_(v); return d.length >= 5 && d.length <= 10 && /^[\d\s-]+$/.test(String(v)); };
  const first = rows[0];
  const hasHeader = first.some(c => Object.keys(RX).some(k => RX[k].test(c))) || !first.some(looksId);
  const body = hasHeader ? rows.slice(1) : rows;
  const find = rx => hasHeader ? first.findIndex(c => rx.test(c)) : -1;

  let idCol = find(RX.id);
  if (idCol < 0) {
    let best = -1, bestShare = 0;
    for (let c = 0; c < ncol; c++) {
      const share = body.filter(r => looksId(r[c] || '')).length / Math.max(1, body.length);
      if (share > bestShare) { best = c; bestShare = share; }
    }
    if (bestShare >= 0.6) idCol = best;
  }
  if (idCol < 0) return parseLoose_(lines, RX);


  const fullCol = find(RX.full), firstCol = find(RX.first), lastCol = find(RX.last), emailCol = find(RX.email);
  const list = [];
  let skipped = 0;
  body.forEach(r => {
    const raw = String(r[idCol] || '').replace(/\D/g, '');
    if (!looksId(r[idCol] || '')) { skipped++; return; }
    let name = '';
    if (fullCol >= 0) name = r[fullCol] || '';
    else if (firstCol >= 0 || lastCol >= 0) name = [r[firstCol] || '', r[lastCol] || ''].join(' ');
    else name = r.filter((c, i) => i !== idCol && i !== emailCol && c && !/@/.test(c) && !/^\d+$/.test(c)).slice(0, 2).join(' ');
    list.push({ rawId: raw, name: name.replace(/\s+/g, ' ').trim() });
  });
  // רוב השורות לא התאימו לעמודות: כנראה רשימה שהוקלדה ביד
  if (list.length < skipped) return parseLoose_(lines, RX);
  return { list: list, skipped: skipped };
}

function RX_ROSTER_() {
  return {
    id: /ת["״׳']?ז|תעודת|זהות|מספר סטודנט|id ?number|idnumber|student ?id/i,
    first: /שם פרטי|first ?name|given/i,
    last: /שם משפחה|surname|last ?name|family/i,
    full: /^שם$|שם מלא|full ?name|^name$/i,
    email: /mail|דוא/i
  };
}

// רשימה שהוקלדה ביד, למשל "123456789 דנה לוי": מוצאים בכל שורה את המספר, והשאר הוא השם
function parseLoose_(lines, RX) {
  const list = [];
  let skipped = 0;
  lines.forEach(l => {
    const m = l.match(/(?:^|[^\d])(\d[\d-]{3,11}\d)(?![\d])/);
    const digits = m ? m[1].replace(/\D/g, '') : '';
    if (digits.length < 5 || digits.length > 10) {
      if (!Object.keys(RX).some(k => RX[k].test(l))) skipped++;   // שורת כותרת לא נספרת כדילוג
      return;
    }
    const name = l.replace(m[1], ' ').replace(/\S+@\S+/g, ' ').replace(/[\t,;|]+/g, ' ').replace(/\s+/g, ' ').trim();
    list.push({ rawId: digits, name: name });
  });
  if (!list.length) throw E_('noidcol');
  return { list: list, skipped: skipped };
}

function splitLine_(line, delim) {
  if (delim === '\t') return line.split('\t');
  const out = [];
  let cur = '', q = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (q) {
      if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; }
      else if (ch === '"') q = false;
      else cur += ch;
    } else if (ch === '"') q = true;
    else if (ch === delim) { out.push(cur); cur = ''; }
    else cur += ch;
  }
  out.push(cur);
  return out;
}

// ===== דוח =====
function attendanceData_() {
  const sessions = rows_(sheet_(SH.SESSIONS), 5)
    .map(v => ({ s: String(v[0]), date: v[1], label: String(v[2] || ''), rounds: Number(v[3]) || 0 }))
    .filter(x => x.s && x.rounds > 0);
  const seen = {};
  rows_(sheet_(SH.CHECKINS), 8).forEach(v => {
    if (v[7] !== STATUS_OK) return;
    const k = String(v[1]) + '|' + normId_(v[3]);
    (seen[k] = seen[k] || {})[Number(v[2])] = true;
  });
  const caught = {};
  rows_(sheet_(SH.SPOT), 5).forEach(v => {
    if (v[4] === 'לא') caught[String(v[1]) + '|' + normId_(v[2])] = true;
  });
  return { sessions: sessions, seen: seen, caught: caught };
}

function cellFor_(d, se, id) {
  const k = se.s + '|' + id;
  if (d.caught[k]) return 'נתפס';
  const n = d.seen[k] ? Object.keys(d.seen[k]).length : 0;
  if (n >= se.rounds) return '✓';
  if (n === 0) return '';
  return n + '/' + se.rounds;
}

function buildReport_() {
  const students = loadStudents_();
  const d = attendanceData_();
  const dateStr = x => (x instanceof Date) ? Utilities.formatDate(x, CFG.TZ, 'dd/MM') : String(x);
  const header = ['ת"ז', 'שם'].concat(d.sessions.map(x => dateStr(x.date) + (x.label ? ' ' + x.label : '')), ['נוכחויות', 'אחוז']);
  const out = [header];
  const bg = [['#ffffff']];
  students.forEach(st => {
    let full = 0;
    const row = [st.rawId, st.name];
    d.sessions.forEach(se => {
      const c = cellFor_(d, se, st.id);
      if (c === '✓') full++;
      row.push(c);
    });
    const share = d.sessions.length ? full / d.sessions.length : 1;
    row.push(full, d.sessions.length ? Math.round(100 * share) + '%' : '');
    out.push(row);
    bg.push([share < CFG.ATTENDANCE_THRESHOLD ? '#f8d7da' : '#ffffff']);
  });

  const ss = ss_();
  let sh = ss.getSheetByName(SH.REPORT);
  if (!sh) sh = ss.insertSheet(SH.REPORT);
  sh.clear();
  sh.setRightToLeft(true);
  sh.getRange(1, 1, out.length, header.length).setValues(out);
  sh.getRange(1, header.length, out.length, 1).setBackgrounds(bg);
  sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  sh.setFrozenRows(1);
  return d.sessions.length;
}

// ===== מיקום: הכיתה = החציון של המדידות הטובות בסבב =====
function farRows_(rows) {
  const far = new Set();
  const byR = {};
  rows.forEach(x => { (byR[x.r] = byR[x.r] || []).push(x); });
  Object.keys(byR).forEach(r => {
    const withLoc = byR[r].filter(x => x.acc > 0);
    const good = withLoc.filter(x => x.acc <= CFG.GEO_GOOD_ACCURACY);
    if (good.length < CFG.GEO_MIN_POINTS) return;
    const lat = median_(good.map(x => x.lat));
    const lng = median_(good.map(x => x.lng));
    withLoc.forEach(x => {
      if (meters_(x.lat, x.lng, lat, lng) - x.acc > CFG.FAR_METERS) far.add(x.row);
    });
  });
  return far;
}
// מזעור מידע: המיקום נחוץ רק בזמן השיעור. בפתיחת שיעור חדש, הרישומים של שיעורים קודמים
// מקבלים סימון "רחוק מהכיתה" במידת הצורך, וקווי הרוחב והאורך נמחקים.
function scrubLocations_() {
  const sh = sheet_(SH.CHECKINS);
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const vals = sh.getRange(2, 1, n, 11).getValues();
  const rows = vals.map((v, i) => ({
    i: i, row: i + 2, s: String(v[1]), r: Number(v[2]),
    lat: Number(v[8]), lng: Number(v[9]), acc: Number(v[10]) || 0
  })).filter(x => x.acc > 0);
  if (!rows.length) return;
  const bySession = {};
  rows.forEach(x => { (bySession[x.s] = bySession[x.s] || []).push(x); });
  Object.keys(bySession).forEach(k => {
    const far = farRows_(bySession[k]);
    bySession[k].forEach(x => {
      if (far.has(x.row)) {
        const f = String(vals[x.i][6] || '');
        if (f.indexOf(FLAG_FAR) < 0) vals[x.i][6] = f ? f + ', ' + FLAG_FAR : FLAG_FAR;
      }
      vals[x.i][8] = ''; vals[x.i][9] = ''; vals[x.i][10] = '';
    });
  });
  sh.getRange(2, 7, n, 1).setValues(vals.map(v => [v[6]]));
  sh.getRange(2, 9, n, 3).setValues(vals.map(v => [v[8], v[9], v[10]]));
}

function median_(a) {
  a = a.slice().sort((x, y) => x - y);
  const m = Math.floor(a.length / 2);
  return a.length % 2 ? a[m] : (a[m - 1] + a[m]) / 2;
}
function meters_(lat1, lng1, lat2, lng2) {
  const R = 6371000, toR = Math.PI / 180;
  const dLat = (lat2 - lat1) * toR, dLng = (lng2 - lng1) * toR;
  const h = Math.sin(dLat / 2) ** 2 + Math.cos(lat1 * toR) * Math.cos(lat2 * toR) * Math.sin(dLng / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(h));
}

// ===== עזרים =====
function props_() { return PropertiesService.getScriptProperties(); }
function cache_() { return CacheService.getScriptCache(); }
function ss_() {
  const id = props_().getProperty('SHEET_ID');
  if (!id) throw E_('nosetup');
  return SpreadsheetApp.openById(id);
}
function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw E_('nosetup');
  return sh;
}
function baseUrl_() {
  const saved = props_().getProperty('WEBAPP_URL');
  if (saved) return saved;
  const u = String(ScriptApp.getService().getUrl() || '');
  return WEBAPP_RX.test(u) ? u : '';
}

// הכתובת שבקוד על המקרן: הדף החיצוני כשהוא פעיל, אחרת דף הגיבוי של Apps Script
function studentUrl_(a, sl) {
  const q = 's=' + encodeURIComponent(a.s) + '&r=' + a.r + '&c=' + code_(a.s, a.r, sl);
  const key = frontendOn_() ? apiKey_() : '';
  return key ? CFG.FRONTEND_URL + '?' + key + '&' + q : baseUrl_() + '?' + q;
}
function frontendOn_() { return !!CFG.FRONTEND_URL && props_().getProperty('FRONTEND') !== 'off'; }
function apiKey_() {
  const m = String(baseUrl_() || '').match(WEBAPP_RX);
  if (!m) return '';
  return 'k=' + m[2] + (m[1] ? '&d=' + encodeURIComponent(m[1]) : '');
}

// השיעור הפעיל. סבב נסגר לבד אחרי ROUND_MINUTES, ושיעור מיום קודם נסגר לבד
function active_() {
  const v = props_().getProperty('ACTIVE');
  if (!v) return null;
  const a = JSON.parse(v);
  if (a.s.slice(0, 8) !== Utilities.formatDate(new Date(), CFG.TZ, 'yyyyMMdd')) {
    props_().deleteProperty('ACTIVE');
    return null;
  }
  if (a.open && a.openedAt && Date.now() - a.openedAt >= CFG.ROUND_MINUTES * 60000) {
    a.open = false;
    setActive_(a);
  }
  return a;
}
function setActive_(a) { props_().setProperty('ACTIVE', JSON.stringify(a)); }

function setSessionRounds_(s, r) {
  const sh = sheet_(SH.SESSIONS);
  const ids = colValues_(sh, 1).map(String);
  const i = ids.lastIndexOf(s);
  if (i >= 0) sh.getRange(i + 2, 4).setValue(r);
}

function secret_() {
  const s = props_().getProperty('SECRET');
  if (!s) throw E_('nosetup');
  return s;
}
function hex_(bytes) {
  return bytes.map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}
function hmac_(msg) { return hex_(Utilities.computeHmacSha256Signature(msg, secret_())); }
function pinHash_(pin) {
  return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, secret_() + '|pin|' + pin));
}
function safeEq_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  let d = 0;
  for (let i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}

function slot_() { return Math.floor(Date.now() / (CFG.ROTATE_SECONDS * 1000)); }
function code_(s, r, slot) { return hmac_(['qr', s, r, slot].join('|')).slice(0, 12); }
function passFor_(s, r, issued) { return hmac_(['pass', s, r, issued].join('|')); }

function verifyPass_(req) {
  req = req || {};
  const issued = Number(req.issued);
  const age = Date.now() - issued;
  if (!issued || age > CFG.PASS_TTL_SECONDS * 1000 || age < -5000) throw E_('expired');
  const s = String(req.s || '');
  const r = Number(req.r);
  if (!safeEq_(passFor_(s, r, issued), String(req.pass || ''))) throw E_('expired');
  return { s: s, r: r };
}

function normId_(x) { return String(x == null ? '' : x).replace(/\D/g, '').replace(/^0+/, ''); }
function cleanToken_(t) {
  t = String(t || '');
  return /^[a-f0-9]{32}$/.test(t) ? t : '';
}

function rows_(sh, ncols) {
  const n = sh.getLastRow() - 1;
  return n < 1 ? [] : sh.getRange(2, 1, n, ncols).getValues();
}
function colValues_(sh, col) {
  const n = sh.getLastRow() - 1;
  return n < 1 ? [] : sh.getRange(2, col, n, 1).getValues().map(v => v[0]);
}

function loadStudents_() {
  return rows_(sheet_(SH.STUDENTS), 5).map((v, i) => ({
    row: i + 2,
    id: normId_(v[0]),
    rawId: String(v[0]),
    name: String(v[1] || '').trim(),
    token: String(v[2] || ''),
    rebinds: Number(v[4]) || 0
  })).filter(x => x.id);
}

function checkinRows_(s) {
  return rows_(sheet_(SH.CHECKINS), 11)
    .map((v, i) => ({
      row: i + 2, s: String(v[1]), r: Number(v[2]), id: normId_(v[3]), name: String(v[4]),
      flags: String(v[6] || ''), status: v[7],
      lat: Number(v[8]), lng: Number(v[9]), acc: Number(v[10]) || 0
    }))
    .filter(x => x.s === s);
}
function spotRows_(s) {
  return rows_(sheet_(SH.SPOT), 5).filter(v => String(v[1]) === s).map(v => ({ id: normId_(v[2]) }));
}

function alreadyIn_(s, r, id) {
  return !!cache_().get(['ck', s, r, id].join('|')) || alreadyInSheet_(s, r, id);
}
// גיבוי למקרה שהמטמון נמחק: בודק רק את הרישומים האחרונים כדי להישאר מהיר
function alreadyInSheet_(s, r, id) {
  if (cache_().get(['cnt', s, r].join('|')) !== null && cache_().get('warm|' + s + '|' + r)) return false;
  const sh = sheet_(SH.CHECKINS);
  const last = sh.getLastRow();
  const n = Math.min(last - 1, 1500);
  if (n < 1) { cache_().put('warm|' + s + '|' + r, '1', CFG.CACHE_TTL); return false; }
  const vals = sh.getRange(last - n + 1, 2, n, 3).getValues();
  const c = cache_();
  let found = false;
  vals.forEach(v => {
    if (String(v[0]) === s && Number(v[1]) === r) {
      c.put(['ck', s, r, normId_(v[2])].join('|'), '1', CFG.CACHE_TTL);
      if (normId_(v[2]) === id) found = true;
    }
  });
  c.put('warm|' + s + '|' + r, '1', CFG.CACHE_TTL);
  return found;
}

function count_(s, r) {
  const k = ['cnt', s, r].join('|');
  const v = cache_().get(k);
  if (v !== null) return Number(v);
  const n = checkinRows_(s).filter(x => x.r === r).length;
  cache_().put(k, String(n), CFG.CACHE_TTL);
  return n;
}
function flagCount_(s) {
  const v = cache_().get('fl|' + s);
  if (v !== null) return Number(v);
  const n = checkinRows_(s).filter(x => x.flags).length;
  cache_().put('fl|' + s, String(n), CFG.CACHE_TTL);
  return n;
}

function shuffle_(a) {
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const t = a[i]; a[i] = a[j]; a[j] = t;
  }
  return a;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw E_('busy');
  try { return fn(); } finally { lock.releaseLock(); }
}

function E_(code, detail) { return new Error('E:' + code + (detail ? '|' + detail : '')); }

function api_(fn) {
  try { return fn(); }
  catch (e) {
    const m = String(e && e.message || e);
    if (m.indexOf('E:') === 0) {
      const parts = m.slice(2).split('|');
      return { error: parts[0], detail: parts[1] || '' };
    }
    console.error(e);
    return { error: 'server' };
  }
}

function adminApi_(pin, fn) {
  return api_(() => {
    const h = props_().getProperty('ADMIN_PIN_HASH');
    if (!h || !safeEq_(pinHash_(String(pin || '')), h)) {
      Utilities.sleep(1500);   // מאט ניחושים בלי לנעול את המרצה בחוץ
      throw E_('badpin');
    }
    return fn();
  });
}
