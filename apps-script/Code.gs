/**
 * סולם — backend ב-Google Apps Script, צמוד לגיליון Google Sheets בדרייב שלך.
 *
 * התקנה (פעם אחת):
 * 1. צור Google Sheet חדש (למשל "סולם — בחירות 2026").
 * 2. Extensions → Apps Script, הדבק את הקובץ הזה, שמור.
 * 3. Project Settings → Script Properties → הוסף GEMINI_API_KEY
 *    (מפתח חינמי מ-aistudio.google.com/apikey). לחלופין ANTHROPIC_API_KEY.
 * 4. העלה את app.html לדרייב (כל תיקייה, אותו חשבון).
 * 5. Deploy → New deployment → Web app: Execute as "Me", Who has access "Anyone".
 *    אשר את ההרשאות. ה-URL שמתקבל הוא הכתובת של השאלון — זו הכתובת לשלוח לאנשים.
 * 6. בעורך: בחר את הפונקציה setupTrigger ולחץ Run (פעם אחת). היא יוצרת טריגר
 *    שרץ כל 5 דקות, משלים ניתוחים שנכשלו ושולח את מיילי הניתוח שממתינים.
 *
 * אין מה למלא ב-CONFIG.endpoint ואין צורך ב-APP_URL: doGet מגיש את app.html
 * מהדרייב ומזריק לתוכו את הכתובת של עצמו. אחרי כל build_bundle.py — פשוט
 * מחליפים את app.html בדרייב, בלי לגעת ב-Apps Script.
 *
 * ADMIN_KEY — חובה כדי לתייג. מחרוזת אקראית שרק אתה מכיר; מסך #admin מבקש
 *    אותה פעם אחת ושומר במכשיר. בלעדיה השרת דוחה כל כתיבת תיוג.
 *
 * APP_FILE_ID (רשות) — אם יש כמה קבצים בשם app.html בדרייב, הגדר אותו כדי לקבע
 * איזה מהם מוגש. APP_URL (רשות) — רק אם תעדיף להגיש את השאלון מאתר אחר.
 * DATA_URL (רשות) — מאיפה הטריגר טוען את השאלות והתיוג כשהוא בונה ניתוח למשיב
 * ותיק בלי packet. ברירת מחדל: האתר ב-GitHub Pages.
 *
 * הגיליון מקבל את הלשוניות האלה אוטומטית:
 *   tags     — התיוג שלך (party_id, q_id, level, principled, note, link, updated)
 *   answers  — שורה לכל משיב, לפי מייל: ts, updated, name, email, token,
 *              עמודה לכל שאלה (level|P אם קו אדום), answers_json, scores_json, analysis_json,
 *              ובסוף: packet_json, mail_opt, mailed_at, tries
 *   history  — הגרסה הקודמת נשמרת כאן בכל פעם שמשיב מעדכן תשובות
 *   log      — שגיאות
 *
 * המייל הוא המזהה החד-ערכי. אחזור נעשה בטוקן אקראי בלבד, שנשלח למייל —
 * לא מספיק להכיר כתובת מייל כדי לפתוח את התשובות של מישהו.
 *
 * מייל הניתוח נשלח רק למי שסימן בטופס "לשלוח לי את הניתוח למייל" (mail_opt=yes).
 * משיבים ותיקים שלא היה להם ניתוח מקבלים ניתוח מהטריגר, אבל לא מייל.
 */
const MAX_TOKENS = 2200;
const MAX_TRIES = 6;            // כמה פעמים הטריגר מנסה לנתח שורה לפני שהוא מוותר עליה
const BATCH_MS = 4 * 60 * 1000; // ריצת טריגר עוצרת אחרי 4 דקות (המגבלה של Apps Script היא 6)
const DEFAULT_DATA_URL = 'https://samstrix782.github.io/sulam/app/';
const EXTRA_COLS = ['packet_json', 'mail_opt', 'mailed_at', 'tries'];

// מנוע הניתוח נבחר לפי המפתח שקיים ב-Script Properties:
//   GEMINI_API_KEY     → Google AI Studio (שכבה חינמית). aistudio.google.com/apikey
//   ANTHROPIC_API_KEY  → Anthropic
// אם שניהם קיימים — גיימיני קודם, ואנתרופיק רק אם גיימיני נכשל. אם אף אחד לא — אין ניתוח, וכל השאר עובד.
const GEMINI_MODELS = ['gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-flash-latest'];
const ANTHROPIC_MODEL = 'claude-sonnet-4-6';

function sheet_(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(header); sh.setFrozenRows(1); }
  return sh;
}
const out_ = (obj) => ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
const norm_ = (s) => String(s || '').trim().toLowerCase();
const prop_ = (k) => PropertiesService.getScriptProperties().getProperty(k);
const log_ = (msg) => sheet_('log', ['ts', 'error']).appendRow([new Date(), String(msg).slice(0, 2000)]);

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === 'tags') return out_({ tags: readTags_() });
  if (action) return out_({ ok: true });
  return serveApp_(e);   // בלי action — מגישים את השאלון עצמו
}

// מגיש את app.html מהדרייב. מזריק לתוכו את כתובת ה-Web app עצמה, כך שאין
// מה למלא ידנית ב-CONFIG.endpoint, ואת טוקן האחזור אם הגיעו עם ?r=...
function serveApp_(e) {
  const file = appFile_();
  if (!file) {
    return HtmlService.createHtmlOutput('<div dir="rtl" style="font-family:Arial;padding:24px">'
      + '<h2>app.html לא נמצא בדרייב</h2>'
      + '<p>העלה את app.html לדרייב, או הגדר ב-Script Properties את APP_FILE_ID עם מזהה הקובץ.</p></div>');
  }
  let html = file.getBlob().getDataAsString('UTF-8');
  const r = String((e && e.parameter && e.parameter.r) || '').replace(/[^a-zA-Z0-9]/g, '');
  const inject = '<script>window.__ENDPOINT__=' + JSON.stringify(ScriptApp.getService().getUrl())
               + ';window.__RECALL__=' + JSON.stringify(r) + ';</script>\n';
  html = html.replace('<script>', inject + '<script>');   // לפני בלוק ה-__DATA__ הראשון
  return HtmlService.createHtmlOutput(html)
    .setTitle('סולם — שאלון ערכים לבחירות 2026')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function appFile_() {
  const id = prop_('APP_FILE_ID');
  if (id) { try { return DriveApp.getFileById(id); } catch (err) {} }
  const it = DriveApp.getFilesByName('app.html');
  let newest = null;
  while (it.hasNext()) {
    const f = it.next();
    if (!newest || f.getLastUpdated() > newest.getLastUpdated()) newest = f;
  }
  return newest;
}
function appUrl_() { return prop_('APP_URL') || ScriptApp.getService().getUrl(); }
function recallUrl_(token) {
  const base = appUrl_();
  return base + (base.indexOf('?') >= 0 ? '&' : '?') + 'r=' + token;
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ error: 'bad json' }); }
  try {
    if (body.action === 'admin_check') return out_(adminOk_(body) ? { ok: true } : { error: adminError_(body) });
    if (body.action === 'tag' || body.action === 'tags_bulk') {
      if (!adminOk_(body)) return out_({ error: adminError_(body) });
      if (body.action === 'tag') return out_(saveTag_(body));
      (body.tags || []).forEach(saveTag_);
      return out_({ ok: true, n: (body.tags || []).length });
    }
    if (body.action === 'submit') return out_(submit_(body));
    if (body.action === 'recall') return out_(recall_(body));
    if (body.action === 'recall_link') return out_(recallLink_(body));
    return out_({ error: 'unknown action' });
  } catch (err) {
    log_(err && err.stack || err);
    return out_({ error: String(err) });
  }
}

/* ---------- tags ---------- */
// כתיבת תיוג היא פעולה של המחבר בלבד. ה-Web app פרוס כ-Anyone והכתובת שלו
// גלויה בדף, אז בלי המפתח הזה כל אחד היה יכול לדרוס את התיוג בגיליון.
// נועל כברירת מחדל: בלי ADMIN_KEY מוגדר בשרת, שום כתיבה לא עוברת.
function adminOk_(b) {
  const k = prop_('ADMIN_KEY');
  return !!k && String(b && b.key || '') === k;
}
function adminError_(b) {
  if (!prop_('ADMIN_KEY')) return 'לא הוגדר ADMIN_KEY ב-Script Properties — התיוג נעול';
  return 'מפתח תיוג שגוי';
}

function readTags_() {
  const sh = sheet_('tags', ['party_id', 'q_id', 'level', 'principled', 'note', 'link', 'updated']);
  const rows = sh.getDataRange().getValues().slice(1);
  const tags = {};
  rows.forEach(r => {
    if (!r[0] || !r[1]) return;
    tags[r[0]] = tags[r[0]] || {};
    tags[r[0]][r[1]] = { level: Number(r[2]) || null, principled: r[3] === true || r[3] === 'TRUE' || r[3] === 1, note: r[4] || '', link: r[5] || '', source: 'sheet' };
  });
  return tags;
}
function saveTag_(t) {
  const sh = sheet_('tags', ['party_id', 'q_id', 'level', 'principled', 'note', 'link', 'updated']);
  const data = sh.getDataRange().getValues();
  const row = [t.party_id, t.q_id, t.level || '', !!t.principled, t.note || '', t.link || '', new Date()];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === t.party_id && data[i][1] === t.q_id) { sh.getRange(i + 1, 1, 1, row.length).setValues([row]); return { ok: true, updated: true }; }
  }
  sh.appendRow(row); return { ok: true, inserted: true };
}

/* ---------- answers sheet ---------- */
const HEAD_FIXED = ['ts', 'updated', 'name', 'email', 'token'];
const TAIL_FIXED = ['answers_json', 'scores_json', 'analysis_json'];

// מחזיר את לשונית answers בפורמט הנוכחי. אם קיימת לשונית ישנה בלי עמודת email,
// היא נשמרת בצד בשם answers_v1 ונפתחת חדשה — כדי שלא נדרוס נתונים.
// העמודות של EXTRA_COLS נוספות בסוף אם חסרות, כך שלשונית קיימת לא צריכה הגירה.
function answersSheet_(qids) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const header = HEAD_FIXED.concat(qids).concat(TAIL_FIXED).concat(EXTRA_COLS);
  let sh = ss.getSheetByName('answers');
  if (sh) {
    const first = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    if (first.indexOf('email') === -1) {
      sh.setName('answers_v1_' + Utilities.formatDate(new Date(), 'UTC', 'yyyyMMdd'));
      sh = null;
    }
  }
  if (!sh) { sh = ss.insertSheet('answers'); sh.appendRow(header); sh.setFrozenRows(1); }
  ensureCols_(sh, EXTRA_COLS);
  return sh;
}
function ensureCols_(sh, names) {
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const missing = names.filter(n => head.indexOf(n) === -1);
  if (missing.length) sh.getRange(1, head.length + 1, 1, missing.length).setValues([missing]);
}
function colIndex_(sh, name) {
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  return head.indexOf(name) + 1; // 1-based, 0 אם לא נמצא
}
function findRowBy_(sh, colName, value) {
  const c = colIndex_(sh, colName);
  if (!c || sh.getLastRow() < 2) return 0;
  const vals = sh.getRange(2, c, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (norm_(vals[i][0]) === norm_(value)) return i + 2;
  return 0;
}
function setCells_(sh, rowNum, obj) {
  Object.keys(obj).forEach(name => {
    const c = colIndex_(sh, name);
    if (c) sh.getRange(rowNum, c).setValue(obj[name]);
  });
}

/* ---------- submit ---------- */
function submit_(b) {
  const email = norm_(b.email);
  if (!email || email.indexOf('@') < 1) return { error: 'צריך כתובת מייל' };
  const qids = b.qids || Object.keys(b.answers || {});
  const sh = answersSheet_(qids);
  const wantsMail = !!b.mail;
  let token, analysis, rowNum;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    rowNum = findRowBy_(sh, 'email', email);
    token = (rowNum ? sh.getRange(rowNum, colIndex_(sh, 'token')).getValue() : '') || newToken_();
    if (rowNum) archive_(sh, rowNum);

    analysis = b.packet ? analyze_(b.packet) : '';
    const cells = qids.map(id => { const a = b.answers[id]; return a && a.level ? (a.level + (a.principled ? '|P' : '')) : ''; });
    const now = new Date();
    const created = rowNum ? sh.getRange(rowNum, 1).getValue() : now;
    const row = [created, now, b.name || '', email, token]
      .concat(cells)
      .concat([JSON.stringify(b.answers || {}), JSON.stringify(b.scores || {}), analysis]);

    if (rowNum) sh.getRange(rowNum, 1, 1, row.length).setValues([row]);
    else { sh.appendRow(row); rowNum = sh.getLastRow(); }
    // תשובות חדשות = ניתוח חדש: מאפסים את סימון המייל והניסיונות, כדי שהמייל הבא ישקף אותן
    setCells_(sh, rowNum, { packet_json: b.packet ? JSON.stringify(b.packet) : '', mail_opt: wantsMail ? 'yes' : '', mailed_at: '', tries: analysis ? 1 : 0 });
  } finally {
    lock.releaseLock();
  }
  // המייל נשלח מחוץ לנעילה, כדי שמשיב אחר לא ימתין לו
  const parsed = parseAnalysis_(analysis);
  if (wantsMail && parsed && sendAnalysisMail_(email, b.name || '', token, parsed)) {
    setCells_(sh, rowNum, { mailed_at: new Date() });
  }
  return { ok: true, token: token, analysis: parsed, mail_pending: wantsMail && !parsed };
}
function newToken_() { return Utilities.getUuid().replace(/-/g, ''); }
function archive_(sh, rowNum) {
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const hist = sheet_('history', ['archived_at'].concat(head));
  const row = sh.getRange(rowNum, 1, 1, sh.getLastColumn()).getValues()[0];
  hist.appendRow([new Date()].concat(row));
}

/* ---------- recall ---------- */
function recall_(b) {
  const token = String(b.token || '').replace(/[^a-zA-Z0-9]/g, '');
  if (token.length < 8) return { error: 'טוקן לא תקין' };
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('answers');
  if (!sh) return { error: 'אין עדיין תשובות' };
  const rowNum = findRowBy_(sh, 'token', token);
  if (!rowNum) return { error: 'לא נמצא' };
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = sh.getRange(rowNum, 1, 1, sh.getLastColumn()).getValues()[0];
  const get = (name) => { const i = head.indexOf(name); return i < 0 ? '' : row[i]; };
  let answers = {}; try { answers = JSON.parse(get('answers_json') || '{}'); } catch (e) {}
  return { ok: true, name: get('name'), email: get('email'),
           answers: answers, analysis: parseAnalysis_(get('analysis_json')) };
}

// שולח קישור אישי למייל. תמיד מחזיר ok — כדי שלא ניתן יהיה לגלות מהתשובה
// אם כתובת מסוימת מסרה טופס או לא.
function recallLink_(b) {
  const email = norm_(b.email);
  if (!email || email.indexOf('@') < 1) return { error: 'צריך כתובת מייל' };
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('answers');
    if (sh) {
      const rowNum = findRowBy_(sh, 'email', email);
      if (rowNum) {
        const token = sh.getRange(rowNum, colIndex_(sh, 'token')).getValue();
        const name = sh.getRange(rowNum, colIndex_(sh, 'name')).getValue();
        const link = recallUrl_(token);
        MailApp.sendEmail({
          to: email,
          subject: 'סולם — הקישור למגש שלך',
          htmlBody: '<div dir="rtl" style="font-family:Arial,sans-serif;font-size:15px;line-height:1.6;color:#131A2A">'
            + '<p>' + (name ? escapeHtml_(String(name)) + ',' : 'שלום,') + '</p>'
            + '<p>הקישור הזה פותח את המגש, הפתקים והניתוח שנכתב לך — ואת 104 התשובות, פתוחות לשינוי:</p>'
            + '<p><a href="' + link + '" style="color:#A81E14">' + link + '</a></p>'
            + '<p style="color:#6E6858;font-size:13px">הקישור אישי. מי שמחזיק בו רואה את התשובות שלך, אז עדיף לא להעביר אותו.<br>'
            + 'אם לא ביקשת את המייל הזה — אפשר להתעלם ממנו.</p></div>'
        });
      }
    }
  } catch (err) {
    log_('recall_link: ' + (err && err.stack || err));
  }
  return { ok: true };
}
function escapeHtml_(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

/* ---------- analysis mail ---------- */
// בונה ושולח את מייל הניתוח. מחזיר true אם נשלח. לא זורק — כישלון נרשם ב-log
// והטריגר ינסה שוב בריצה הבאה.
function sendAnalysisMail_(email, name, token, analysis) {
  try {
    if (MailApp.getRemainingDailyQuota() < 1) { log_('mail quota exhausted, deferred: ' + email); return false; }
    const link = recallUrl_(token);
    MailApp.sendEmail({
      to: email,
      subject: 'סולם — מפת הפשרות שלך',
      htmlBody: analysisMailHtml_(name, analysis, link),
      body: analysisMailText_(name, analysis, link),
      name: 'סולם'
    });
    return true;
  } catch (err) {
    log_('analysis mail: ' + (err && err.stack || err));
    return false;
  }
}

// מייל בטבלאות ובסגנון inline — כך הוא נראה אותו דבר ב-Gmail, באאוטלוק ובטלפון.
// הצבעים והמבנה כמו מסך "מפת הפשרות" באפליקציה.
function analysisMailHtml_(name, a, link) {
  const C = { ink: '#131A2A', soft: '#4A5166', muted: '#6E6858', paper: '#FBFAF7', slip: '#FFFDF6', ground: '#EFEDE6', rule: '#DBD6C9', line: '#A81E14', assent: '#2F5D50' };
  const serif = '\'Frank Ruhl Libre\',Georgia,\'Times New Roman\',serif';
  const sans = 'Heebo,Arial,sans-serif';
  const p = (t, st) => '<p style="margin:0 0 10px;' + (st || '') + '">' + t + '</p>';
  const list = (items, color, mark) => (items || []).map(t =>
    '<tr><td dir="rtl" style="padding:3px 0;font-size:14px;line-height:1.55;color:' + C.ink + ';text-align:right">'
    + '<span style="color:' + color + ';font-weight:700">' + mark + '</span>&nbsp; ' + escapeHtml_(t) + '</td></tr>').join('');

  let body = '';
  if (typeof a === 'object' && a && Array.isArray(a.parties)) {
    if (a.profile) body += '<tr><td style="padding:4px 24px 18px;font-family:' + serif + ';font-size:18px;line-height:1.45;color:' + C.ink + '">' + escapeHtml_(a.profile) + '</td></tr>';
    a.parties.forEach(pp => {
      body += '<tr><td style="padding:0 16px 14px">'
        + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:' + C.slip + ';border:1px solid ' + C.rule + '">'
        + '<tr><td style="padding:16px 16px 6px">'
        +   '<table role="presentation" cellpadding="0" cellspacing="0"><tr>'
        +     '<td style="width:52px;height:60px;border:1.5px solid ' + C.ink + ';background:#fff;text-align:center;font-family:' + serif + ';font-weight:800;font-size:20px;color:' + C.ink + '">' + escapeHtml_(pp.letters || '—') + '</td>'
        +     '<td style="padding-right:12px">'
        +       '<div style="font-size:16px;font-weight:700;color:' + C.ink + '">' + escapeHtml_(pp.name || '') + '</div>'
        +       (pp.match != null ? '<div style="font-size:13px;color:' + C.muted + ';margin-top:2px">' + escapeHtml_(pp.match) + '% התאמה</div>' : '')
        +     '</td></tr></table>'
        + '</td></tr>'
        + '<tr><td style="padding:10px 16px 0"><div style="font-size:12px;font-weight:700;letter-spacing:.03em;color:' + C.assent + '">למה כן</div>'
        +   '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' + list(pp.pros, C.assent, '✓') + '</table></td></tr>'
        + '<tr><td style="padding:12px 16px 0"><div style="font-size:12px;font-weight:700;letter-spacing:.03em;color:' + C.line + '">למה לא</div>'
        +   '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' + list(pp.cons, C.line, '✕') + '</table></td></tr>'
        + (pp.tradeoff ? '<tr><td style="padding:12px 16px 0"><div style="border-right:3px solid ' + C.ink + ';padding:4px 10px 4px 0;font-family:' + serif + ';font-size:15px;line-height:1.5;color:' + C.soft + '">' + escapeHtml_(pp.tradeoff) + '</div></td></tr>' : '')
        + '<tr><td style="height:16px"></td></tr>'
        + '</table></td></tr>';
    });
  } else {
    String(a || '').split(/\n{2,}/).forEach(par => {
      body += '<tr><td style="padding:0 24px 10px;font-size:14px;line-height:1.6;color:' + C.ink + '">' + escapeHtml_(par).replace(/\n/g, '<br>') + '</td></tr>';
    });
  }

  return '<!doctype html><html lang="he" dir="rtl"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>'
    + '<body style="margin:0;padding:0;background:' + C.ground + '">'
    + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:' + C.ground + '"><tr><td align="center" style="padding:20px 8px">'
    + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" dir="rtl" style="max-width:560px;background:' + C.paper + ';font-family:' + sans + ';color:' + C.ink + ';text-align:right">'
    + '<tr><td style="padding:22px 24px 6px;border-top:4px solid ' + C.ink + '">'
    +   '<div style="font-family:' + serif + ';font-size:24px;font-weight:800">סולם</div>'
    +   '<div style="font-size:13px;color:' + C.muted + ';margin-top:2px">מפת הפשרות שלך · בחירות 2026</div></td></tr>'
    + '<tr><td style="padding:14px 24px 6px;font-size:15px;line-height:1.6">'
    +   p((name ? escapeHtml_(name) + ',' : 'שלום,'))
    +   p('כאן הרשימות שהכי קרובות לתשובות שלך — ומה כל אחת מהן דורשת ממך לוותר עליו.')
    + '</td></tr>'
    + body
    + '<tr><td style="padding:8px 24px 22px">'
    +   '<a href="' + link + '" style="display:block;background:' + C.ink + ';color:' + C.paper + ';text-decoration:none;text-align:center;padding:14px 10px;font-size:15px;font-weight:700">לפתוח את המגש המלא ולשנות תשובות</a>'
    + '</td></tr>'
    + '<tr><td style="padding:14px 24px 22px;border-top:1px solid ' + C.rule + ';font-size:12px;line-height:1.6;color:' + C.muted + '">'
    +   p('הניתוח נכתב על ידי מודל שפה, לפי התשובות שלך ולפי תיוג העמדות של המפלגות בסולם. הוא לא המלצת הצבעה — ההחלטה שלך.', 'font-size:12px')
    +   p('קיבלת את המייל כי ביקשת אותו בטופס. הקישור למעלה אישי: מי שמחזיק בו רואה את התשובות שלך, אז עדיף לא להעביר את המייל הלאה.', 'font-size:12px;margin:0')
    + '</td></tr>'
    + '</table></td></tr></table></body></html>';
}
function analysisMailText_(name, a, link) {
  const lines = [(name ? name + ',' : 'שלום,'), ''];
  if (typeof a === 'object' && a && Array.isArray(a.parties)) {
    if (a.profile) lines.push(a.profile, '');
    a.parties.forEach(pp => {
      lines.push((pp.letters ? pp.letters + ' · ' : '') + (pp.name || '') + (pp.match != null ? ' — ' + pp.match + '% התאמה' : ''));
      (pp.pros || []).forEach(t => lines.push('  ✓ ' + t));
      (pp.cons || []).forEach(t => lines.push('  ✕ ' + t));
      if (pp.tradeoff) lines.push('  ' + pp.tradeoff);
      lines.push('');
    });
  } else lines.push(String(a || ''), '');
  lines.push('המגש המלא: ' + link, '', 'הניתוח נכתב על ידי מודל שפה ואינו המלצת הצבעה.');
  return lines.join('\n');
}

/* ---------- background: retries, backfill, mail ---------- */
// להרצה ידנית פעם אחת מהעורך. יוצר טריגר שמריץ את processPending כל 5 דקות.
function setupTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processPending')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processPending').timeBased().everyMinutes(5).create();
  Logger.log('נוצר טריגר: processPending כל 5 דקות. מכסת מייל שנותרה היום: ' + MailApp.getRemainingDailyQuota());
}

// עובר על לשונית answers:
// 1. שורה בלי ניתוח — מנתח (לפי packet_json, ובמשיבים ותיקים בלי packet — בונה אותו מהתשובות).
// 2. שורה עם ניתוח, mail_opt=yes ובלי mailed_at — שולח את המייל.
// עוצר אחרי BATCH_MS, וממשיך בריצה הבאה.
function processPending() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;   // ריצה קודמת עדיין עובדת
  const started = Date.now();
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('answers');
    if (!sh || sh.getLastRow() < 2) return;
    ensureCols_(sh, EXTRA_COLS);
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getValues();
    const ix = (n) => head.indexOf(n);
    let data = null;   // נטען רק אם צריך לבנות packet

    for (let i = 0; i < rows.length; i++) {
      if (Date.now() - started > BATCH_MS) break;
      const r = rows[i], rowNum = i + 2;
      const email = norm_(r[ix('email')]);
      if (!email) continue;
      let analysis = String(r[ix('analysis_json')] || '');
      const tries = Number(r[ix('tries')]) || 0;

      if (!analysis && tries < MAX_TRIES) {
        let packet = null;
        try { packet = JSON.parse(r[ix('packet_json')] || 'null'); } catch (e) {}
        if (!packet) {
          let answers = {}; try { answers = JSON.parse(r[ix('answers_json')] || '{}'); } catch (e) {}
          if (!Object.keys(answers).length) continue;
          data = data || loadData_();
          packet = buildPacket_(data, answers);
          sh.getRange(rowNum, ix('packet_json') + 1).setValue(JSON.stringify(packet));
        }
        analysis = analyze_(packet);
        sh.getRange(rowNum, ix('tries') + 1).setValue(tries + 1);
        if (!analysis) continue;
        sh.getRange(rowNum, ix('analysis_json') + 1).setValue(analysis);
      }

      if (analysis && r[ix('mail_opt')] === 'yes' && !r[ix('mailed_at')]) {
        const parsed = parseAnalysis_(analysis);
        if (parsed && sendAnalysisMail_(email, r[ix('name')], r[ix('token')], parsed)) {
          sh.getRange(rowNum, ix('mailed_at') + 1).setValue(new Date());
        }
      }
    }
  } catch (err) {
    log_('processPending: ' + (err && err.stack || err));
  } finally {
    lock.releaseLock();
  }
}

/* ---------- packet (server side) ---------- */
// השאלות והתיוג כפי שהם באתר, עם התיוג המאושר מהגיליון מעליהם — בדיוק כמו load() באפליקציה.
function loadData_() {
  const base = prop_('DATA_URL') || DEFAULT_DATA_URL;
  const q = JSON.parse(UrlFetchApp.fetch(base + 'questions.json').getContentText());
  const p = JSON.parse(UrlFetchApp.fetch(base + 'parties.json').getContentText());
  const sheetTags = readTags_();
  p.parties.forEach(party => { if (sheetTags[party.id]) party.tags = Object.assign({}, party.tags, sheetTags[party.id]); });
  return { q: q, p: p };
}

// העתק של score() ו-submitAnswers() מ-app/index.html. משמש רק למשיבים שאין להם
// packet_json שמור (מסרו לפני שהעמודה נוספה) — שינוי בחישוב באפליקציה צריך לעבור גם לכאן.
const TOP_N = 5;
function scoreParty_(values, answers, party) {
  let sw = 0, sd = 0;
  const byValue = {}, redlines = [], theirs = [];
  values.forEach(v => {
    let vw = 0, vd = 0, vn = 0;
    v.questions.forEach(q => {
      const a = answers[q.id], t = party.tags[q.id];
      if (!a || !a.level || !t || !t.level) return;
      const d = Math.abs(a.level - t.level), w = a.principled ? 2 : 1, s = 1 - d / 4;
      sw += w; sd += w * s; vw += w; vd += w * s; vn++;
      if (a.principled && d >= 2) redlines.push({ q: q, v: v, a: a, t: t });
      else if (t.principled && d >= 2) theirs.push({ q: q, v: v, a: a, t: t });
    });
    byValue[v.id] = vn ? { pct: Math.round(100 * vd / vw) } : null;
  });
  return { pct: sw ? Math.round(100 * sd / sw) : null, byValue: byValue, redlines: redlines, theirs: theirs };
}
function buildPacket_(data, answers) {
  const values = data.q.values;
  const qs = [];
  values.forEach(v => v.questions.forEach(q => qs.push({ q: q, v: v })));
  const rung = (q, level) => level + ' — ' + ((q.rungs.find(x => x.level === level) || {}).text || '');
  const ranked = data.p.parties.map(p => ({ p: p, s: scoreParty_(values, answers, p) }))
    .filter(r => r.s.pct !== null)
    .sort((a, b) => b.s.pct - a.s.pct);
  return {
    profile: {
      answered: Object.keys(answers).filter(k => answers[k] && answers[k].level).length,
      total: qs.length,
      redlines_by_value: Object.fromEntries(values.map(v => [v.name, v.questions.filter(q => answers[q.id] && answers[q.id].principled).length]))
    },
    parties: ranked.slice(0, TOP_N).map(({ p, s }) => ({
      name: p.name, letters: p.letters || '', leader: p.leader, match_pct: s.pct, components: p.components,
      user_redlines_crossed: s.redlines.map(r => ({ value: r.v.name, question: r.q.title, user: rung(r.q, r.a.level), party: rung(r.q, r.t.level), party_note: r.t.note, estimate: r.t.source === 'estimate' })),
      party_principles_user_far_from: s.theirs.map(r => ({ value: r.v.name, question: r.q.title, user: r.a.level, party: rung(r.q, r.t.level), party_note: r.t.note, estimate: r.t.source === 'estimate' })),
      agreements: qs.filter(({ q }) => {
        const a = answers[q.id], t = p.tags[q.id];
        return a && a.level && t && t.level && Math.abs(a.level - t.level) <= 1 && (a.principled || t.principled);
      }).map(({ q, v }) => ({ value: v.name, question: q.title, user: answers[q.id].level, party: p.tags[q.id].level,
        user_redline: !!answers[q.id].principled, party_principle: !!p.tags[q.id].principled, party_note: p.tags[q.id].note, estimate: p.tags[q.id].source === 'estimate' })),
      by_value: Object.fromEntries(values.map(v => [v.name, s.byValue[v.id] ? s.byValue[v.id].pct : null]))
    }))
  };
}

/* ---------- analysis ---------- */
const SYSTEM = `אתה מנתח שאלון ערכים לבחירות לכנסת ה-26. המשתמש ענה על סולמות של 1–5 בשבעה ערכים וסימן לכל שאלה אם היא "קו אדום" בשבילו או פתוחה לדיון. המפלגות תויגו על אותם סולמות בשאלון, לפי הצבעות, מצעים, ראיונות והצהרות של הרשימות ומועמדיהן, וגם אצלן סומן מה עקרוני ומה סחיר. התיוג הוא נתון, לא שלך — אל תוסיף ידע חיצוני על המפלגות ואל תתקן אותו; כתוב "לפי התיוג" כשזה רלוונטי.

חלק מהתיוג הוא הערכה שטרם אושרה: פריט כזה מסומן בחומר ב-"estimate": true. נקודה שנשענת על פריט כזה — כתוב בה "לפי הערכה" ואל תציג אותה כעובדה. אם רוב מה שמבדיל רשימה מסוים נשען על הערכות, אמור זאת במשפט הפשרה שלה.

הטקסט מוצג למשתמש במסך ונשלח אליו במייל: פנה אליו ישירות, בגוף שני. אל תזכיר את מחבר השאלון או אנשים אחרים בשמם, חוץ ממועמדי הרשימות כפי שהם בחומר.

המשימה: להחזיר 2–3 מפלגות אפשריות למשתמש, לא אחת. בחר לפי ההתאמה הכללית, אבל בעיקר לפי מה שסומן כמשמעותי משני הצדדים: קו אדום של המשתמש שהמפלגה חוצה הוא סיבה כבדה נגד; עיקרון של המפלגה שהמשתמש קרוב אליו הוא סיבה כבדה בעד. אם המפלגה עם ההתאמה המספרית הגבוהה ביותר חוצה קו אדום, אמור זאת במפורש ואל תדחוק אותה — תן למשתמש להחליט.

החזר JSON בלבד, בלי טקסט לפניו או אחריו, בלי גדרות קוד, במבנה הזה בדיוק:
{
  "profile": "משפט אחד שמסכם את הפרופיל של המשתמש — באילו ערכים סימן הכי הרבה קווים אדומים ומה זה אומר על הפשרות שמחכות לו",
  "parties": [
    {
      "name": "שם הרשימה כפי שהוא בחומר",
      "letters": "אותיות הרשימה כפי שהן בחומר",
      "match": 84,
      "pros": ["2–3 נקודות קונקרטיות: באילו שאלות המשתמש והמפלגה על אותו שלב או שלב סמוך, במיוחד היכן שזה עקרוני לאחד מהם"],
      "cons": ["2–3 נקודות: קווים אדומים שנחצים, פערים גדולים, ועקרונות של המפלגה שהמשתמש רחוק מהם ויצטרך לחיות איתם"],
      "tradeoff": "משפט אחד או שניים — מה המשתמש מקבל ברשימה הזאת ומה הוא מוותר עליו"
    }
  ]
}

כללים: עברית, ישיר, בלי חנופה ובלי "כל המפלגות לגיטימיות". השתמש בשמות השאלות ובניסוח השלבים כפי שהם בחומר. אל תמציא עמדות שלא בתיוג. אל תגיד למשתמש למי להצביע — הצג את הפשרה. כל נקודה ב-pros וב-cons היא משפט אחד או שניים, לא יותר. עד 450 מילים בסך הכול.`;

// גיימיני קודם; אם הוא נכשל ויש מפתח אנתרופיק — אנתרופיק. מחזיר '' אם שניהם נכשלו,
// והטריגר ינסה שוב בריצה הבאה.
function analyze_(packet) {
  const gem = prop_('GEMINI_API_KEY'), ant = prop_('ANTHROPIC_API_KEY');
  let out = '';
  if (gem) out = analyzeGemini_(packet, gem);
  if (!out && ant) out = analyzeAnthropic_(packet, ant);
  return out;
}

// שמות הדגמים אצל גוגל משתנים מדי כמה חודשים, אז מנסים כמה לפי הסדר
// ומפסיקים בראשון שעונה. GEMINI_MODEL ב-Script Properties גובר על הרשימה.
// 503/429/500 (עומס) — עוברים לדגם הבא, שיש לו קיבולת נפרדת, במקום לוותר.
function analyzeGemini_(packet, key) {
  const pinned = prop_('GEMINI_MODEL');
  const models = (pinned ? [pinned] : []).concat(GEMINI_MODELS);
  const body = {
    system_instruction: { parts: [{ text: SYSTEM }] },
    contents: [{ role: 'user', parts: [{ text: JSON.stringify(packet) }] }],
    generationConfig: { temperature: 0.4, maxOutputTokens: MAX_TOKENS, responseMimeType: 'application/json' }
  };
  const errs = [];
  for (let i = 0; i < models.length; i++) {
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/'
              + encodeURIComponent(models[i]) + ':generateContent';
    const res = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-goog-api-key': key },
      payload: JSON.stringify(body)
    });
    const code = res.getResponseCode();
    if (code === 200) {
      try {
        const j = JSON.parse(res.getContentText());
        const parts = j.candidates[0].content.parts || [];
        return stripFence_(parts.map(pt => pt.text || '').join(''));
      } catch (e) { errs.push('gemini parse: ' + String(e)); break; }
    }
    errs.push('gemini ' + models[i] + ' ' + code + ': ' + res.getContentText().slice(0, 200));
    if ([400, 404, 429, 500, 503].indexOf(code) === -1) break;   // 401/403 = מפתח שגוי; אין טעם בדגם הבא
    if (code === 429 || code === 503) Utilities.sleep(1500);
  }
  log_(errs.join(' | '));
  return '';
}

function analyzeAnthropic_(packet, key) {
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model: ANTHROPIC_MODEL, max_tokens: MAX_TOKENS, system: SYSTEM, messages: [{ role: 'user', content: JSON.stringify(packet) }] })
  });
  if (res.getResponseCode() !== 200) {
    log_('anthropic ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 500));
    return '';
  }
  try {
    return stripFence_(JSON.parse(res.getContentText()).content.map(c => c.text || '').join(''));
  } catch (e) { return ''; }
}

// להרצה ידנית מהעורך (Run) כדי לראות איזה מנוע פעיל ואיזה דגם עונה.
// התוצאה נכתבת ל-Execution log, לא לגיליון.
function pingModel() {
  const gem = prop_('GEMINI_API_KEY'), ant = prop_('ANTHROPIC_API_KEY');
  Logger.log('GEMINI_API_KEY: ' + (gem ? 'מוגדר' : 'חסר') + ' | ANTHROPIC_API_KEY: ' + (ant ? 'מוגדר' : 'חסר'));
  if (!gem && !ant) { Logger.log('אין מפתח — לא ייכתב ניתוח, כל השאר עובד'); return; }
  const out = analyze_({ profile: { answered: 1, total: 1 }, parties: [{ name: 'בדיקה', letters: 'בד', match_pct: 50, user_redlines_crossed: [], party_principles_user_far_from: [], agreements: [], by_value: {} }] });
  Logger.log(out ? ('עובד. תשובה: ' + String(out).slice(0, 300)) : 'נכשל — הפירוט בלשונית log בגיליון');
}

// להרצה ידנית: שולח לעצמך (לבעלים של הסקריפט) מייל ניתוח לדוגמה, כדי לראות איך הוא נראה.
function previewMail() {
  const me = Session.getEffectiveUser().getEmail();
  const sample = { profile: 'רוב הקווים האדומים שלך בערך "מדינה בטוחה" — הפשרות הקשות יהיו שם.',
    parties: [{ name: 'רשימה לדוגמה', letters: 'דג', match: 81, pros: ['אתם על אותו שלב בגיוס חרדים, וזה קו אדום אצלך.'],
      cons: ['הרשימה חוצה את הקו האדום שלך בשאלת מינוי שופטים (לפי התיוג).'], tradeoff: 'מקבלים התאמה בביטחון, מוותרים במשפט.' }] };
  MailApp.sendEmail({ to: me, subject: 'סולם — תצוגה מקדימה של מייל הניתוח', htmlBody: analysisMailHtml_('', sample, appUrl_()), body: analysisMailText_('', sample, appUrl_()), name: 'סולם' });
  Logger.log('נשלח אל ' + me);
}

function stripFence_(t) {
  const s = String(t || '').trim();
  const m = /^```(?:json)?\s*([\s\S]*?)\s*```$/.exec(s);
  return m ? m[1].trim() : s;
}
// מחזיר אובייקט אם הניתוח הוא JSON תקין במבנה הצפוי, אחרת את הטקסט כמו שהוא.
// כך המסך יודע להציג מבנה כשיש, ולנפול לפסקאות רגילות כשאין.
function parseAnalysis_(raw) {
  const s = stripFence_(raw);
  if (!s) return '';
  try {
    const o = JSON.parse(s);
    if (o && Array.isArray(o.parties)) return o;
  } catch (e) {}
  return s;
}
