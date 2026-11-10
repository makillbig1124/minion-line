/**
 * Minion family car schedule — backend on Google Sheets.
 * Deploy as Web app (Execute as: Me, Who has access: Anyone).
 * Script property LINE_CHANNEL_ID = the Channel ID of your LINE Login channel.
 * Every request must carry a LINE ID token; the script verifies it with LINE.
 */
// LINE Login Channel ID (a public number, not a secret). Script property LINE_CHANNEL_ID overrides it.
const LINE_CHANNEL_ID = '2012518451';
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
  // Daily 07:00 summary to the family group (only sends when there are bookings that day).
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'dailySummary')) {
    ScriptApp.newTrigger('dailySummary').timeBased().atHour(7).everyDays(1).inTimezone('Asia/Bangkok').create();
  }
}

/** Creates the sheets the first time the app runs, so setup() never has to be run by hand. */
function ensureReady_() {
  if (!SpreadsheetApp.getActive().getSheetByName(MEMBERS)) setup();
}

function doGet() {
  ensureReady_();
  const hasChannel = !!(PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ID') || LINE_CHANNEL_ID);
  return ContentService.createTextOutput('Minion API is ready. Sheets OK. LINE channel: ' + (hasChannel ? 'set' : 'NOT set yet'));
}

function doPost(e) {
  let out;
  try {
    ensureReady_();
    const req = JSON.parse(e.postData.contents);
    if (Array.isArray(req.events)) {           // LINE Messaging API webhook (the family-group bot)
      handleWebhook_(req.events);
      return ContentService.createTextOutput('OK');
    }
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
  const channelId = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ID') || LINE_CHANNEL_ID;
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

// =====================================================================
// LINE bot for the family group (Messaging API).
// Script property LINE_ACCESS_TOKEN = long-lived channel access token of the LINE OA.
// Replies (free) when someone types a keyword; pushes a 07:00 summary on days with bookings.
// =====================================================================
const LIFF_URL = 'https://liff.line.me/2012518451-2GtFIIPh';
const BOT_KEYWORDS = ['จอง', 'จองรถ', 'ตาราง', 'ตารางรถ', 'คิว', 'คิวรถ', 'ใครใช้รถ', 'minion', 'มินเนี่ยน'];
const TH_DAYS = ['อา.', 'จ.', 'อ.', 'พ.', 'พฤ.', 'ศ.', 'ส.'];
const TH_MONTHS = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];

function handleWebhook_(events) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_ACCESS_TOKEN');
  if (!token) return;
  events.forEach(ev => {
    try {
      if (ev.source && ev.source.type === 'group') {
        PropertiesService.getScriptProperties().setProperty('GROUP_ID', ev.source.groupId);
      }
      if (ev.type === 'join') {
        lineApi_(token, 'reply', { replyToken: ev.replyToken, messages: [
          { type: 'text', text: 'สวัสดีครับ 🚗 Minion มาแล้ว!\nพิมพ์ "จอง" หรือ "ตาราง" เพื่อดูคิวรถและเปิดตารางจอง\nทุกเช้า 7 โมงจะสรุปคิวรถของวันนั้นให้ (ถ้ามีคนจอง)' },
          scheduleFlex_('ตารางรถ Minion', 2)
        ] });
      } else if (ev.type === 'message' && ev.message.type === 'text') {
        const text = ev.message.text.trim().toLowerCase();
        if (BOT_KEYWORDS.indexOf(text) !== -1) {
          lineApi_(token, 'reply', { replyToken: ev.replyToken, messages: [scheduleFlex_('ตารางรถ Minion', 2)] });
        }
      }
    } catch (err) {
      console.error(err);
    }
  });
}

/** Time trigger (07:00 Asia/Bangkok). Run it by hand from the editor to test. */
function dailySummary() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('LINE_ACCESS_TOKEN'), groupId = props.getProperty('GROUP_ID');
  if (!token || !groupId) return;
  const today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  if (!read_(sheet_(BOOKINGS), B_COLS).some(b => b.date === today)) return;   // quiet on free days, saves quota
  lineApi_(token, 'push', { to: groupId, messages: [scheduleFlex_('สรุปคิวรถวันนี้', 1)] });
}

function scheduleFlex_(title, days) {
  const members = read_(sheet_(MEMBERS), M_COLS);
  const bookings = read_(sheet_(BOOKINGS), B_COLS);
  const body = [];
  for (let i = 0; i < days; i++) {
    const d = new Date(Date.now() + i * 86400000);
    const iso = Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd');
    const parts = iso.split('-').map(Number);
    const local = new Date(parts[0], parts[1] - 1, parts[2]);
    const label = (i === 0 ? 'วันนี้' : 'พรุ่งนี้') + ' · ' + TH_DAYS[local.getDay()] + ' ' + parts[2] + ' ' + TH_MONTHS[parts[1] - 1];
    body.push({ type: 'text', text: label, weight: 'bold', size: 'sm', color: '#1D3F73', margin: i ? 'lg' : 'none' });
    const list = bookings.filter(b => b.date === iso).sort((a, b) => toMin_(a.start) - toMin_(b.start));
    if (!list.length) body.push({ type: 'text', text: 'ว่างทั้งวัน ✅', size: 'sm', color: '#1E7A4C', margin: 'sm' });
    list.forEach(b => {
      const m = members.find(x => x.id === b.memberId);
      body.push({ type: 'box', layout: 'baseline', spacing: 'md', margin: 'sm', contents: [
        { type: 'text', text: b.start + '–' + b.end, size: 'sm', color: '#5F6A80', flex: 0 },
        { type: 'text', text: (m ? m.name : 'ไม่ระบุ') + (b.purpose ? ' · ' + b.purpose : ''), size: 'sm', wrap: true,
          color: (m && /^#[0-9a-fA-F]{6}$/.test(m.color)) ? m.color : '#1F2A44', weight: 'bold' }
      ] });
    });
  }
  return {
    type: 'flex', altText: title + ' · กดเพื่อเปิดตารางจอง',
    contents: { type: 'bubble', size: 'kilo',
      header: { type: 'box', layout: 'vertical', backgroundColor: '#FFD60A', paddingAll: '14px',
        contents: [{ type: 'text', text: '🚗 ' + title, weight: 'bold', size: 'md', color: '#1F2A44' }] },
      body: { type: 'box', layout: 'vertical', contents: body },
      footer: { type: 'box', layout: 'vertical', contents: [
        { type: 'button', style: 'primary', color: '#1D3F73', height: 'sm', action: { type: 'uri', label: 'เปิดตารางจอง', uri: LIFF_URL } }
      ] } }
  };
}

function lineApi_(token, kind, payload) {
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/' + kind, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token }, payload: JSON.stringify(payload)
  });
  if (res.getResponseCode() !== 200) console.error('LINE ' + kind + ' failed: ' + res.getContentText());
}
