/**
 * EMR COMPLIANCE DASHBOARD
 * Developed by: Bluemoonways | info.bluemoonways@gmail.com
 * Version 1.1 | October 2026
 *
 * Backend (Code.gs)   [Speed-Optimized + Roles + CO Regional Access]
 *
 * Roles:
 *   - Admin: FULL RIGHTS. Sees all divisions; add / edit / delete queries (even after answered);
 *            write or change any answer (even after verified); Mark Resolved / Return; all reports.
 *   - CO: Add / Edit / Delete queries (before answered) ONLY for divisions in their own Province.
 *   - Senior: Reviews answers: Verified / Unverified (returned with a reason).
 *   - Division: Sees only its own queries, writes answers.
 *   - Viewer: Read-only access to all divisions and all reports.
 *
 * v1.1 changes:
 *   - api_addBulkQueries: section validation, duplicate skip, max 60 per call, returns { added, skipped }
 *   - api_setStatus: optional expectedOn check (stops Senior verifying an answer that changed meanwhile)
 *   - withLock_: friendly "server busy" message
 *   - auth_: constant-time password compare, failed-attempt counter reset on success
 */
const DB = {
  USERS: 'Users',
  QUERIES: 'Queries',
  SECTIONS: 'Sections',
  DIV_SHEET: 'Sheet1',
  DIV_FIRST_ROW: 4,
  TEMPLATE: 'Main',
  TEMPLATE_FIRST_ROW: 6,
  TZ: 'Asia/Karachi',
  QH: ['ID', 'Month', 'Division', 'Section', 'Query', 'Answer', 'Status', 'Query Date', 'Added By',
       'Answered By', 'Answered On', 'Verified By', 'Verified On', 'Senior Remarks', 'Return Count'],
  UH: ['Name', 'Role', 'Division', 'User ID', 'Password', 'Email'],
  VERIFIED: 'Verified',
  UNVERIFIED: 'Unverified',
  ROLES: ['Senior', 'Division', 'Viewer', 'Admin', 'CO']
};

const Q = { ID: 0, MONTH: 1, DIV: 2, SECTION: 3, QUERY: 4, ANSWER: 5, STATUS: 6, QDATE: 7, ADDEDBY: 8,
            ABY: 9, AON: 10, VBY: 11, VON: 12, REMARKS: 13, RETCOUNT: 14 };

const CK = { USERS: 'users_v2', DIVS: 'divs_v2', SECS: 'secs_v2' };
const TTL = { USERS: 120, DIVS: 300, SECS: 300 };
const SESSION_DAYS = 7;
const MEMO = {};

const DEFAULT_SECTIONS = [
  'Fund Position & Purposes', 'Bank Reconciliation (BRS)', 'Wajiba & Shari Process', 'DM ERP Reconciliation',
  'Fund Transfers', 'Contra Accounts', 'AR/AP from Division/HO Accounts', 'Customers Report', 'Suppliers Report',
  'Item Donation', 'FC, Gold, Silver Usher Sales', 'MIS Data Accuracy', 'E Expense', 'E Raseed', 'Rent Booking',
  'Account Movement', 'Consolidated Financials'
];

const DEFAULT_DIVISIONS = [
  ['Punjab', 'Rawalpindi'], ['Punjab', 'Gujrat'], ['Punjab', 'Gujranwala'], ['Punjab', 'Sheikhupura'],
  ['Punjab', 'Faisalabad'], ['Punjab', 'Multan'], ['Punjab', 'Sargodha'], ['Punjab', 'Sahiwal'],
  ['Punjab', 'DG Khan'], ['Punjab', 'Bahawalpur'], ['Islamabad', 'Islamabad'], ['Lahore', 'Lahore'],
  ['Karachi', 'Karachi'], ['KPK', 'Hazara'], ['KPK', 'Peshawar'], ['KPK', 'Dera Ismail Khan'],
  ['Kashmir', 'Mirpur'], ['Kashmir', 'Poonch'], ['Kashmir', 'Muzaffarabad'], ['Gilgit', 'Gilgit'],
  ['Sindh', 'Hyderabad'], ['Sindh', 'Nawabshah'], ['Sindh', 'Larkana'], ['Sindh', 'Sukkur'],
  ['Sindh', 'Bhanbhore'], ['Sindh', 'Jacobabad'], ['Sindh', 'Mirpur Khas'], ['Balochistan', 'Balochistan']
].map(r => ({ province: r[0], division: r[1] }));

/* ===================================================================
 * MENU (inside the Google Sheet)
 * =================================================================== */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('EMR Dashboard')
    .addItem('1. Setup sheets', 'setupSheets')
    .addItem('2. Generate missing passwords', 'generatePins')
    .addItem('3. Install daily reminder', 'installReminder')
    .addSeparator()
    .addItem('4. Add missing CO users', 'addMissingCOUsers')
    .addSeparator()
    .addItem('Send reminders now', 'sendRemindersNow')
    .addItem('Show dashboard link', 'showLink')
    .addItem('Clear cache (apply sheet changes now)', 'clearCacheMenu')
    .addItem('Sign out everyone (reset sessions)', 'resetSessions')
    .addToUi();
}

function clearCacheMenu() {
  clearCaches_();
  SpreadsheetApp.getUi().alert('Cache cleared. Your sheet changes will now apply immediately.');
}

function resetSessions() {
  PropertiesService.getScriptProperties().deleteProperty('TOKEN_SECRET');
  delete MEMO.secret;
  SpreadsheetApp.getUi().alert('All users have been signed out. Everyone must now sign in again with their User ID and password.');
}

/**
 * Adds any missing CO users to the existing Users tab. Existing users are left untouched.
 */
function addMissingCOUsers() {
  const ss = ss_();
  const us = ss.getSheetByName(DB.USERS);
  if (!us || us.getLastRow() < 1) {
    SpreadsheetApp.getUi().alert('Users tab not found. Please run Setup sheets first.');
    return;
  }

  const existing = us.getLastRow() >= 2
    ? us.getRange(2, 1, us.getLastRow() - 1, DB.UH.length).getValues()
    : [];

  const haveCO = {};
  existing.forEach(r => {
    const role = String(r[1] || '').trim();
    const prov = String(r[2] || '').trim();
    if (role === 'CO' && prov) haveCO[prov] = true;
  });

  const usedIds = {};
  existing.forEach(r => {
    const id = String(r[3] || '').trim().toLowerCase();
    if (id) usedIds[id] = true;
  });

  const divs = getDivisions_();
  const provinces = {};
  divs.forEach(d => { if (d.province) provinces[d.province] = true; });

  const rows = [];
  Object.keys(provinces).sort().forEach(p => {
    if (haveCO[p]) return;

    let uid = p.toLowerCase().replace(/[^a-z0-9]/g, '') + '_co';
    if (usedIds[uid]) {
      let n = 2;
      while (usedIds[uid + n]) n++;
      uid = uid + n;
    }
    usedIds[uid] = true;

    rows.push([p + ' CO', 'CO', p, uid, randomPin_(), '']);
  });

  if (!rows.length) {
    clearCaches_();
    SpreadsheetApp.getUi().alert('A CO already exists for every province. Nothing was added.');
    return;
  }

  us.getRange(us.getLastRow() + 1, 1, rows.length, DB.UH.length)
    .setNumberFormat('@')
    .setValues(rows);

  clearCaches_();
  SpreadsheetApp.getUi().alert(
    rows.length + ' new CO user(s) added!\n\nOpen the Users tab and note down the passwords.'
  );
}

function setupSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());

    let qs = ss.getSheetByName(DB.QUERIES);
    if (!qs) qs = ss.insertSheet(DB.QUERIES);
    if (qs.getMaxColumns() < DB.QH.length) qs.insertColumnsAfter(qs.getMaxColumns(), DB.QH.length - qs.getMaxColumns());
    if (qs.getLastRow() < 1) {
      qs.getRange(1, 1, 1, DB.QH.length).setNumberFormat('@').setValues([DB.QH]).setFontWeight('bold').setBackground('#e8eaed');
      qs.setFrozenRows(1);
    }
    qs.getRange(1, 1, qs.getMaxRows(), DB.QH.length).setNumberFormat('@');

    delete MEMO.ss; delete MEMO.qs;
    qsheet_();

    let sec = ss.getSheetByName(DB.SECTIONS);
    if (!sec) sec = ss.insertSheet(DB.SECTIONS);
    if (sec.getLastRow() < 2) {
      sec.getRange('A1').setValue('Section').setFontWeight('bold');
      let names = [];
      const main = ss.getSheetByName(DB.TEMPLATE);
      if (main && main.getLastRow() >= DB.TEMPLATE_FIRST_ROW) {
        names = main.getRange(DB.TEMPLATE_FIRST_ROW, 2, main.getLastRow() - DB.TEMPLATE_FIRST_ROW + 1, 1)
          .getValues().map(r => String(r[0]).trim()).filter(Boolean);
      }
      if (!names.length) names = DEFAULT_SECTIONS;
      sec.getRange(2, 1, names.length, 1).setValues(names.map(n => [n]));
      sec.setColumnWidth(1, 320);
    }

    let us = ss.getSheetByName(DB.USERS);
    if (!us) us = ss.insertSheet(DB.USERS);
    if (us.getLastRow() < 2) {
      us.getRange(1, 1, 1, DB.UH.length).setNumberFormat('@').setValues([DB.UH]).setFontWeight('bold').setBackground('#e8eaed');
      us.setFrozenRows(1);

      const rows = [
        ['Admin User', 'Admin', '', 'admin', randomPin_(), ''],
        ['Senior User', 'Senior', '', 'senior', randomPin_(), ''],
        ['Viewer User', 'Viewer', '', 'viewer', randomPin_(), '']
      ];

      const divs = getDivisions_();
      const provinces = {};
      divs.forEach(d => { provinces[d.province] = true; });
      Object.keys(provinces).forEach(p => {
        rows.push([p + ' CO', 'CO', p, p.toLowerCase().replace(/[^a-z0-9]/g, '') + '_co', randomPin_(), '']);
      });

      divs.forEach(d => {
        rows.push([d.division + ' Division', 'Division', d.division,
                   d.division.toLowerCase().replace(/[^a-z0-9]/g, ''), randomPin_(), '']);
      });

      us.getRange(2, 1, rows.length, DB.UH.length).setNumberFormat('@').setValues(rows);
      us.setColumnWidths(1, 6, 150);
    }

    clearCaches_();
    SpreadsheetApp.getUi().alert('Setup complete.');
  } catch (err) {
    SpreadsheetApp.getUi().alert('Setup failed:\n\n' + err.message);
    throw err;
  }
}

function generatePins() {
  const us = ss_().getSheetByName(DB.USERS);
  const n = us.getLastRow() - 1;
  let filled = 0;
  if (n > 0) {
    const pins = us.getRange(2, 5, n, 1).getValues();
    pins.forEach(r => { if (!String(r[0]).trim()) { r[0] = randomPin_(); filled++; } });
    us.getRange(2, 5, n, 1).setNumberFormat('@').setValues(pins);
  }
  clearCaches_();
  SpreadsheetApp.getUi().alert(filled + ' new password(s) generated.');
}

function showLink() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert(url ? 'Dashboard link:\n\n' + url : 'Not deployed yet.');
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('EMR Compliance')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getDivisionsForUser_(u) {
  const all = getDivisions_();
  if (['Senior', 'Admin', 'Viewer'].indexOf(u.role) >= 0) return all;
  if (u.role === 'Division') return all.filter(d => d.division === u.division);
  if (u.role === 'CO') return all.filter(d => d.province === u.division);
  return [];
}

function api_login(a) {
  const u = auth_(a);
  const token = (a && a.token) ? '' : makeToken_(u.id, s_(a.pin).trim());

  const allowedDivsObj = getDivisionsForUser_(u);
  const allowedDivNames = allowedDivsObj.map(d => d.division);

  const all = readQueries_();
  const now = Utilities.formatDate(new Date(), DB.TZ, 'yyyy-MM');
  const months = {};
  months[now] = true;
  all.forEach(r => { if (r[Q.MONTH]) months[r[Q.MONTH]] = true; });
  const list = Object.keys(months).sort().reverse();
  const month = list.indexOf(now) >= 0 ? now : list[0];

  return {
    user: { name: u.name, role: u.role, division: u.division },
    token: token,
    sections: getSections_(),
    divisions: allowedDivsObj,
    months: list,
    month: month,
    data: all.filter(r => r[Q.MONTH] === month && allowedDivNames.indexOf(r[Q.DIV]) >= 0).map(obj_)
  };
}

/**
 * User changes their own password. Needs the current password.
 * Sessions are signed with the password, so after a change every OTHER device is signed out
 * automatically; this device gets a fresh token in the reply.
 */
function api_changePin(a, oldPin, newPin) {
  const u = auth_(a);
  oldPin = s_(oldPin).trim();
  newPin = s_(newPin).trim();
  if (!oldPin) throw new Error('Enter your current password.');
  if (!/^\S{6,20}$/.test(newPin)) throw new Error('New password must be 6 to 20 characters (letters, numbers or symbols, no spaces).');
  if (newPin.charAt(0) === "'") throw new Error('New password cannot start with an apostrophe.');
  if (newPin === oldPin) throw new Error('New password must be different from the current password.');

  const cache = CacheService.getScriptCache();
  const key = 'fail_' + u.id.toLowerCase();
  if (Number(cache.get(key) || 0) >= 5) throw new Error('Too many wrong attempts. Try again in 10 minutes.');

  return withLock_(() => {
    const sh = ss_().getSheetByName(DB.USERS);
    const n = sh ? sh.getLastRow() - 1 : 0;
    if (n < 1) throw new Error('User not found.');
    const rows = sh.getRange(2, 1, n, DB.UH.length).getValues();
    for (let i = 0; i < n; i++) {
      if (s_(rows[i][3]).trim().toLowerCase() === u.id.toLowerCase()) {
        if (!safeEq_(s_(rows[i][4]).trim(), oldPin)) {
          cache.put(key, String(Number(cache.get(key) || 0) + 1), 600);
          throw new Error('Current password is wrong.');
        }
        sh.getRange(i + 2, 5).setNumberFormat('@').setValue(newPin);
        clearCaches_();
        cache.remove(key);
        return { token: makeToken_(u.id, newPin) };
      }
    }
    throw new Error('User not found.');
  });
}

function api_overview(a, month) {
  const u = auth_(a);
  const allowedDivsObj = getDivisionsForUser_(u);
  const allowedDivNames = allowedDivsObj.map(d => d.division);

  const rows = readQueries_().filter(r => r[Q.MONTH] === month && allowedDivNames.indexOf(r[Q.DIV]) >= 0);

  const provinces = {};
  allowedDivsObj.forEach(d => { provinces[d.division] = d.province; });

  const by = {};
  const totals = blank_();
  rows.forEach(r => {
    const s = stage_(r);
    const d = r[Q.DIV];
    by[d] = by[d] || blank_();
    by[d].total++; by[d][s]++;
    totals.total++; totals[s]++;
  });

  let divisions = Object.keys(provinces).map(d => Object.assign({ division: d, province: provinces[d] }, by[d] || blank_()));
  return { month: month, totals: totals, divisions: divisions };
}

function api_list(a, month) {
  const u = auth_(a);
  const allowedDivNames = getDivisionsForUser_(u).map(d => d.division);
  return readQueries_()
    .filter(r => r[Q.MONTH] === month && allowedDivNames.indexOf(r[Q.DIV]) >= 0)
    .map(obj_);
}

/**
 * Adds many queries at once (one per section).
 * - Only valid sections are accepted.
 * - Exact duplicates (same month + division + section + text) are skipped, never overwritten.
 * - Existing queries are never changed or deleted.
 * Returns { added, skipped }.
 */
function api_addBulkQueries(a, p) {
  const u = auth_(a);
  need_(u, ['Admin', 'CO']);

  p = p || {};
  const month = String(p.month || '');
  const div = String(p.division || '');
  const queries = Array.isArray(p.queries) ? p.queries : [];
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) throw new Error('Month format must be YYYY-MM.');
  if (!div) throw new Error('Choose a division.');
  if (!queries.length) throw new Error('Write at least one query.');
  if (queries.length > 60) throw new Error('Too many queries in one go.');

  const allowed = getDivisionsForUser_(u).map(d => d.division);
  if (allowed.indexOf(div) < 0) throw new Error('You cannot add queries to this division.');
  const sections = getSections_();

  return withLock_(() => {
    const existing = readQueries_();
    const dup = {};
    existing.forEach(r => {
      if (r[Q.MONTH] === month) dup[r[Q.DIV] + '|' + r[Q.SECTION] + '|' + r[Q.QUERY]] = true;
    });
    const stamp = now_();
    const rows = [];
    let skipped = 0;
    queries.forEach(q => {
      q = q || {};
      const section = String(q.section || '');
      const text = String(q.text || '').trim().slice(0, 2000);
      if (!text) return;
      const key = div + '|' + section + '|' + text;
      if (sections.indexOf(section) < 0 || dup[key]) { skipped++; return; }
      dup[key] = true;
      const r = new Array(DB.QH.length).fill('');
      r[Q.ID] = newId_(); r[Q.MONTH] = month; r[Q.DIV] = div; r[Q.SECTION] = section; r[Q.QUERY] = text;
      r[Q.STATUS] = DB.UNVERIFIED; r[Q.QDATE] = stamp; r[Q.ADDEDBY] = u.name; r[Q.RETCOUNT] = '0';
      rows.push(r);
    });
    if (rows.length) {
      const sh = qsheet_();
      const need = sh.getLastRow() + rows.length;
      if (need > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows() + 200);
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, DB.QH.length).setNumberFormat('@').setValues(rows);
    }
    return { added: rows.length, skipped: skipped };
  });
}

function api_editQuery(a, id, text) {
  const u = auth_(a);
  need_(u, ['Admin', 'CO']);

  text = String(text || '').trim();
  if (!text) throw new Error('Query cannot be empty.');
  if (text.length > 2000) throw new Error('Query is too long.');

  return withLock_(() => {
    const f = find_(id);
    // Admin can edit even after an answer; CO only before an answer
    if (u.role !== 'Admin' && String(f.v[Q.ANSWER]).trim()) {
      throw new Error('This query already has an answer, so it cannot be edited.');
    }

    const allowed = getDivisionsForUser_(u).map(d => d.division);
    if (allowed.indexOf(f.v[Q.DIV]) < 0) throw new Error('You cannot edit this query.');

    f.v[Q.QUERY] = text;
    save_(f);
    return obj_(f.v);
  });
}

function api_deleteQuery(a, id) {
  const u = auth_(a);
  need_(u, ['Admin', 'CO']);

  return withLock_(() => {
    const f = find_(id);
    // Admin can delete even after an answer; CO only before an answer
    if (u.role !== 'Admin' && String(f.v[Q.ANSWER]).trim()) {
      throw new Error('This query already has an answer, so it cannot be deleted.');
    }

    const allowed = getDivisionsForUser_(u).map(d => d.division);
    if (allowed.indexOf(f.v[Q.DIV]) < 0) throw new Error('You cannot delete this query.');

    qsheet_().deleteRow(f.row);
    return { deleted: id };
  });
}

function api_saveAnswer(a, id, answer) {
  const u = auth_(a);
  need_(u, ['Division', 'Admin']);
  answer = String(answer || '').trim();
  if (answer.length > 5000) throw new Error('Answer is too long.');
  return withLock_(() => {
    const f = find_(id);
    if (u.role !== 'Admin' && f.v[Q.DIV] !== u.division) throw new Error('This query belongs to another division.');
    if (u.role !== 'Admin' && f.v[Q.STATUS] === DB.VERIFIED) throw new Error('Already verified.');
    f.v[Q.ANSWER] = answer;
    f.v[Q.STATUS] = DB.UNVERIFIED;
    f.v[Q.ABY] = answer ? u.name : '';
    f.v[Q.AON] = answer ? now_() : '';
    f.v[Q.VBY] = ''; f.v[Q.VON] = ''; f.v[Q.REMARKS] = '';
    save_(f);
    return obj_(f.v);
  });
}

/**
 * Senior / Admin: Mark Resolved (status = Verified) or Return (status = Unverified + remarks).
 * expectedOn (optional): the "Answered On" stamp the reviewer was looking at. If the answer was
 * changed in the meantime, the action is refused so a new answer is never verified unseen.
 */
function api_setStatus(a, id, status, remarks, expectedOn) {
  const u = auth_(a);
  need_(u, ['Senior', 'Admin']);
  remarks = String(remarks || '').trim();
  return withLock_(() => {
    const f = find_(id);
    if (expectedOn != null && String(expectedOn) !== f.v[Q.AON]) {
      throw new Error('Division ne is dauran answer badal diya hai. Page refresh karke naya answer dekhein.');
    }
    if (status === DB.VERIFIED) {
      if (!String(f.v[Q.ANSWER]).trim()) throw new Error('There is no answer to verify yet.');
      f.v[Q.STATUS] = DB.VERIFIED;
      f.v[Q.VBY] = u.name; f.v[Q.VON] = now_(); f.v[Q.REMARKS] = '';
    } else if (status === DB.UNVERIFIED) {
      if (!remarks) throw new Error('Write a short reason for returning this answer.');
      f.v[Q.STATUS] = DB.UNVERIFIED;
      f.v[Q.VBY] = ''; f.v[Q.VON] = ''; f.v[Q.REMARKS] = remarks.slice(0, 1000);
      f.v[Q.RETCOUNT] = String((Number(f.v[Q.RETCOUNT]) || 0) + 1);
    } else {
      throw new Error('Unknown status.');
    }
    save_(f);
    return obj_(f.v);
  });
}

function installReminder() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'dailyReminder') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyReminder').timeBased().everyDays(1).atHour(9).inTimezone(DB.TZ).create();
  SpreadsheetApp.getUi().alert('Daily reminder (9 AM) installed.');
}

function dailyReminder() { sendReminders_(); }
function sendRemindersNow() {
  const r = sendReminders_();
  SpreadsheetApp.getUi().alert('Division emails: ' + r.divisions + '\nSenior emails: ' + r.seniors);
}

function sendReminders_() {
  const link = ScriptApp.getService().getUrl() || '(dashboard link)';
  const need = {}, review = { n: 0 };
  readQueries_().forEach(r => {
    const s = stage_(r);
    if (s === 'awaiting_answer' || s === 'returned') need[r[Q.DIV]] = (need[r[Q.DIV]] || 0) + 1;
    if (s === 'awaiting_review') review.n++;
  });
  const out = { divisions: 0, seniors: 0 };
  users_(true).forEach(u => {
    if (!u.email) return;
    if (u.role === 'Division' && need[u.division]) {
      MailApp.sendEmail(u.email, '[EMR Compliance] ' + need[u.division] + ' queries need your answer',
        'Hello,\n\n' + need[u.division] + ' audit queries are waiting for your answer.\n\nDashboard: ' + link + '\n');
      out.divisions++;
    }
    if (u.role === 'Senior' && review.n) {
      MailApp.sendEmail(u.email, '[EMR Compliance] ' + review.n + ' answers waiting for review',
        'Hello,\n\n' + review.n + ' answers are waiting for your review.\n\nDashboard: ' + link + '\n');
      out.seniors++;
    }
  });
  return out;
}

function ss_() {
  if (MEMO.ss) return MEMO.ss;
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  MEMO.ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  return MEMO.ss;
}

function qsheet_() {
  if (MEMO.qs) return MEMO.qs;
  const sh = ss_().getSheetByName(DB.QUERIES);
  if (!sh) return null;
  try {
    const last = DB.QH.length;
    if (sh.getMaxColumns() < last) sh.insertColumnsAfter(sh.getMaxColumns(), last - sh.getMaxColumns());
    if (sh.getLastRow() >= 1) {
      const cur = String(sh.getRange(1, last).getValue()).trim();
      if (cur !== DB.QH[last - 1]) {
        sh.getRange(1, last, sh.getMaxRows(), 1).setNumberFormat('@');
        sh.getRange(1, last).setValue(DB.QH[last - 1]).setFontWeight('bold').setBackground('#e8eaed');
      }
    }
  } catch (e) {}
  MEMO.qs = sh;
  return sh;
}

function cached_(key, ttl, loader) {
  if (MEMO[key]) return MEMO[key];
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  if (hit) {
    try { MEMO[key] = JSON.parse(hit); return MEMO[key]; } catch (e) {}
  }
  const value = loader();
  MEMO[key] = value;
  try { cache.put(key, JSON.stringify(value), ttl); } catch (e) {}
  return value;
}

function clearCaches_() {
  try { CacheService.getScriptCache().removeAll([CK.USERS, CK.DIVS, CK.SECS]); } catch (e) {}
  Object.keys(MEMO).forEach(k => { delete MEMO[k]; });
}

function randomPin_() { return String(Math.floor(100000 + Math.random() * 900000)); }
function now_() { return Utilities.formatDate(new Date(), DB.TZ, 'dd MMM yyyy, HH:mm'); }
function newId_() { return 'Q' + Utilities.getUuid().replace(/-/g, '').slice(0, 9).toUpperCase(); }
function s_(x) { return x == null ? '' : String(x); }
function blank_() { return { total: 0, awaiting_answer: 0, returned: 0, awaiting_review: 0, verified: 0 }; }

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); }
  catch (e) { throw new Error('Server busy hai, thori der baad dobara try karein.'); }
  try { return fn(); } finally { lock.releaseLock(); }
}

function readUsers_() {
  const sh = ss_().getSheetByName(DB.USERS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, DB.UH.length).getValues().map(r => ({
    name: s_(r[0]).trim(), role: s_(r[1]).trim(), division: s_(r[2]).trim(),
    id: s_(r[3]).trim(), pin: s_(r[4]).trim(), email: s_(r[5]).trim()
  }));
}

function users_(fresh) {
  if (!fresh) return cached_(CK.USERS, TTL.USERS, readUsers_);
  const list = readUsers_();
  MEMO[CK.USERS] = list;
  try { CacheService.getScriptCache().put(CK.USERS, JSON.stringify(list), TTL.USERS); } catch (e) {}
  return list;
}

function auth_(a) {
  if (a && a.token) return userFromToken_(a.token);
  if (!a || !a.id || !a.pin) throw new Error('Please sign in.');
  const id = s_(a.id).trim().toLowerCase();
  const pin = s_(a.pin).trim();
  const cache = CacheService.getScriptCache();
  const key = 'fail_' + id;
  if (Number(cache.get(key) || 0) >= 5) throw new Error('Too many wrong attempts. Try again in 10 minutes.');

  const find = list => list.find(x => x.id.toLowerCase() === id);
  const bad = x => !x || !x.pin || !safeEq_(x.pin, pin) || DB.ROLES.indexOf(x.role) < 0;

  let u = find(users_());
  if (bad(u)) u = find(users_(true));
  if (bad(u)) {
    cache.put(key, String(Number(cache.get(key) || 0) + 1), 600);
    throw new Error('User ID or password is wrong.');
  }
  cache.remove(key);
  return { name: u.name || u.id, role: u.role, division: u.division, id: u.id };
}

function secret_() {
  if (MEMO.secret) return MEMO.secret;
  const props = PropertiesService.getScriptProperties();
  let sec = props.getProperty('TOKEN_SECRET');
  if (!sec) {
    sec = Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('TOKEN_SECRET', sec);
  }
  MEMO.secret = sec;
  return sec;
}

function hmacHex_(msg) {
  return Utilities.computeHmacSha256Signature(msg, secret_())
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function sig_(payload, pin) {
  return hmacHex_(payload + '|' + hmacHex_('pin|' + pin).slice(0, 16));
}

function safeEq_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  let r = 0;
  for (let i = 0; i < a.length; i++) r |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return r === 0;
}

function makeToken_(id, pin) {
  const payload = id + '|' + (Date.now() + SESSION_DAYS * 86400000);
  return Utilities.base64EncodeWebSafe(payload).replace(/=+$/, '') + '.' + sig_(payload, pin);
}

function userFromToken_(token) {
  const bad = new Error('Session expired. Please sign in again.');
  const parts = String(token || '').split('.');
  if (parts.length !== 2 || !parts[1]) throw bad;
  let payload;
  try {
    let b = parts[0];
    while (b.length % 4) b += '=';
    payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(b)).getDataAsString();
  } catch (e) { throw bad; }
  const k = payload.lastIndexOf('|');
  if (k < 1) throw bad;
  const id = payload.slice(0, k).toLowerCase();
  const exp = Number(payload.slice(k + 1));
  if (!exp || exp < Date.now()) throw bad;
  const find = list => list.find(x => x.id.toLowerCase() === id);

  const ok = x => x && x.pin && DB.ROLES.indexOf(x.role) >= 0 && safeEq_(parts[1], sig_(payload, x.pin));

  let u = find(users_());
  if (!ok(u)) u = find(users_(true));
  if (!ok(u)) throw bad;
  return { name: u.name || u.id, role: u.role, division: u.division, id: u.id };
}

function need_(u, roles) {
  if (roles.indexOf(u.role) < 0) throw new Error('Your role (' + u.role + ') cannot do this.');
}

function getDivisions_() {
  return cached_(CK.DIVS, TTL.DIVS, () => {
    const sh = ss_().getSheetByName(DB.DIV_SHEET);
    let list = [];
    if (sh && sh.getLastRow() >= DB.DIV_FIRST_ROW) {
      list = sh.getRange(DB.DIV_FIRST_ROW, 3, sh.getLastRow() - DB.DIV_FIRST_ROW + 1, 2).getValues()
        .map(r => ({ province: s_(r[0]).trim(), division: s_(r[1]).trim() }))
        .filter(d => d.division);
    }
    return list.length ? list : DEFAULT_DIVISIONS.slice();
  });
}

function getSections_() {
  return cached_(CK.SECS, TTL.SECS, () => {
    const sh = ss_().getSheetByName(DB.SECTIONS);
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => s_(r[0]).trim()).filter(Boolean);
  });
}

function readQueries_() {
  const sh = qsheet_();
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, DB.QH.length).getValues()
    .map(r => r.map(s_))
    .filter(r => r[Q.ID]);
}

function find_(id) {
  const sh = qsheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) throw new Error('Query not found.');
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < n; i++) {
    if (s_(ids[i][0]) === id) {
      const v = sh.getRange(i + 2, 1, 1, DB.QH.length).getValues()[0].map(s_);
      return { row: i + 2, v: v };
    }
  }
  throw new Error('Query not found. Refresh the page.');
}

function save_(f) {
  qsheet_().getRange(f.row, 1, 1, DB.QH.length).setNumberFormat('@').setValues([f.v]);
}

function stage_(r) {
  if (r[Q.STATUS] === DB.VERIFIED) return 'verified';
  if (!r[Q.ANSWER].trim()) return 'awaiting_answer';
  return r[Q.REMARKS].trim() ? 'returned' : 'awaiting_review';
}

function obj_(r) {
  return {
    id: r[Q.ID], month: r[Q.MONTH], division: r[Q.DIV], section: r[Q.SECTION], query: r[Q.QUERY],
    answer: r[Q.ANSWER], status: r[Q.STATUS] || DB.UNVERIFIED, queryDate: r[Q.QDATE], addedBy: r[Q.ADDEDBY],
    answeredBy: r[Q.ABY], answeredOn: r[Q.AON], verifiedBy: r[Q.VBY], verifiedOn: r[Q.VON],
    remarks: r[Q.REMARKS], returnCount: Number(r[Q.RETCOUNT]) || 0, stage: stage_(r)
  };
}