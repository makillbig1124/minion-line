/**
 * Minion family car schedule — backend on Google Sheets.
 * Deploy as Web app (Execute as: Me, Who has access: Anyone).
 * Script property LINE_CHANNEL_ID = the Channel ID of your LINE Login channel.
 * Every request must carry a LINE ID token; the script verifies it with LINE.
 */
const MEMBERS = 'Members', BOOKINGS = 'Bookings';
const M_COLS = ['id', 'name', 'color', 'lineUserId', 'addedAt'];
const B_COLS = ['id', 'memberId', 'date', 'start', 'end', 'purpose', 'note', 'createdBy', 'updatedAt'];

/** Run once from the editor: creates the two sheets and the four family drivers. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, MEMBERS, M_COLS);
  ensureSheet_(ss, BOOKINGS, B_COLS);
  const sh = ss.getSheetByName(MEMBERS);
  if (sh.getLastRow() < 2) {
    const now = new Date().toISOString();
    [['big', 'Big', '#2f5c9e'], ['game', 'Game', '#e0a100'], ['yui', 'Yui', '#c2408a'], ['papa', 'Papa', '#1e8a56']]
      .forEach(r => appendRow_(sh, M_COLS, { id: r[0], name: r[1], color: r[2], lineUserId: '', addedAt: now }));
  }
}

function doGet() {
  return ContentService.createTextOutput('Minion API is running.');
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const me = verifyLine_(req.idToken);
    out = handle_(req, me);
  } catch (err) {
    out = { ok: false, code: err.code || 'error', message: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function handle_(req, me) {
  const action = req.action;
  if (action === 'list') return payload_(me);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (action === 'claim') {
      const sh = sheet_(MEMBERS);
      const rows = read_(sh, M_COLS);
      if (rows.some(r => r.lineUserId === me.sub)) fail_('taken', 'This LINE account is already linked');
      const row = rows.find(r => r.id === req.memberId);
      if (!row) fail_('invalid', 'Unknown member');
      if (row.lineUserId) fail_('taken', 'Name already taken');
      writeCell_(sh, M_COLS, row._row, 'lineUserId', me.sub);
      return payload_(me);
    }

    // Everything below needs a linked family member.
    const myMember = read_(sheet_(MEMBERS), M_COLS).find(r => r.lineUserId === me.sub);
    if (!myMember) fail_('not_member', 'Pick your name first');

    if (action === 'saveBooking') {
      const b = req.booking || {};
      const clean = {
        memberId: str_(b.memberId, 40),
        date: match_(b.date, /^\d{4}-\d{2}-\d{2}$/),
        start: match_(b.start, /^([01]\d|2[0-3]):[0-5]\d$/),
        end: match_(b.end, /^(([01]\d|2[0-3]):[0-5]\d|24:00)$/),
        purpose: str_(b.purpose, 80),
        note: str_(b.note, 200),
        createdBy: myMember.id,
        updatedAt: new Date().toISOString()
      };
      if (toMin_(clean.end) <= toMin_(clean.start)) fail_('invalid', 'End must be after start');
      if (!read_(sheet_(MEMBERS), M_COLS).some(r => r.id === clean.memberId)) fail_('invalid', 'Unknown driver');
      const sh = sheet_(BOOKINGS);
      const existing = b.id ? read_(sh, B_COLS).find(r => r.id === b.id) : null;
      if (existing) {
        clean.id = existing.id;
        clean.createdBy = existing.createdBy || myMember.id;
        writeRow_(sh, B_COLS, existing._row, clean);
      } else {
        clean.id = Utilities.getUuid();
        appendRow_(sh, B_COLS, clean);
      }
      return payload_(me);
    }

    if (action === 'deleteBooking') {
      const sh = sheet_(BOOKINGS);
      const row = read_(sh, B_COLS).find(r => r.id === req.id);
      if (row) sh.deleteRow(row._row);
      return payload_(me);
    }

    if (action === 'addMember') {
      appendRow_(sheet_(MEMBERS), M_COLS, {
        id: Utilities.getUuid().slice(0, 8),
        name: str_(req.name, 30) || fail_('invalid', 'Name required'),
        color: match_(req.color, /^#[0-9a-fA-F]{6}$/),
        lineUserId: '',
        addedAt: new Date().toISOString()
      });
      return payload_(me);
    }

    if (action === 'deleteMember') {
      const sh = sheet_(MEMBERS);
      const row = read_(sh, M_COLS).find(r => r.id === req.id);
      if (row) sh.deleteRow(row._row);
      return payload_(me);
    }

    fail_('invalid', 'Unknown action');
  } finally {
    lock.releaseLock();
  }
}

/** What the page receives. People who haven't picked a name only see the names. */
function payload_(me) {
  const members = read_(sheet_(MEMBERS), M_COLS);
  const mine = members.find(r => r.lineUserId === me.sub);
  return {
    ok: true,
    me: { memberId: mine ? mine.id : '', lineName: me.name },
    members: members.map(r => ({ id: r.id, name: r.name, color: r.color || '#8a94a6', linked: !!r.lineUserId })),
    bookings: mine ? read_(sheet_(BOOKINGS), B_COLS).map(r => ({
      id: r.id, memberId: r.memberId, date: r.date, start: r.start, end: r.end, purpose: r.purpose, note: r.note
    })) : []
  };
}

// ---------- LINE ID token check (cached 10 min) ----------
function verifyLine_(token) {
  if (!token) fail_('auth', 'Missing LINE token');
  const channelId = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ID');
  if (!channelId) fail_('config', 'LINE_CHANNEL_ID is not set in Script properties');
  const key = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token));
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  const res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post', payload: { id_token: token, client_id: channelId }, muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) fail_('auth', 'LINE token invalid or expired');
  const j = JSON.parse(res.getContentText());
  const me = { sub: j.sub, name: j.name || '' };
  const ttl = Math.max(1, Math.min(600, Math.floor(j.exp - Date.now() / 1000)));
  cache.put(key, JSON.stringify(me), ttl);
  return me;
}

// ---------- sheet helpers (all cells stored as plain text) ----------
function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) fail_('config', 'Sheet "' + name + '" missing — run setup() once');
  return sh;
}
function ensureSheet_(ss, name, cols) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
  sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@');
  sh.setFrozenRows(1);
}
function read_(sh, cols) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, cols.length).getDisplayValues()
    .map((v, i) => { const o = { _row: i + 2 }; cols.forEach((c, j) => o[c] = v[j]); return o; })
    .filter(o => o.id);
}
function appendRow_(sh, cols, obj) {
  writeRow_(sh, cols, sh.getLastRow() + 1, obj);
}
function writeRow_(sh, cols, row, obj) {
  sh.getRange(row, 1, 1, cols.length).setNumberFormat('@').setValues([cols.map(c => obj[c] == null ? '' : String(obj[c]))]);
}
function writeCell_(sh, cols, row, col, value) {
  sh.getRange(row, cols.indexOf(col) + 1).setNumberFormat('@').setValue(value);
}
function str_(v, max) { return String(v == null ? '' : v).trim().slice(0, max); }
function match_(v, re) { const s = str_(v, 40); if (!re.test(s)) fail_('invalid', 'Bad value: ' + s); return s; }
function toMin_(t) { const p = t.split(':'); return Number(p[0]) * 60 + Number(p[1]); }
function fail_(code, message) { const e = new Error(message); e.code = code; throw e; }
