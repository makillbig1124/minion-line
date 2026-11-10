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
  installTriggers_();
}

/** Daily 07:00 summary + a check every 5 minutes for bookings starting in about 30 minutes. */
function installTriggers_() {
  const have = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  if (have.indexOf('dailySummary') === -1) {
    ScriptApp.newTrigger('dailySummary').timeBased().atHour(7).everyDays(1).inTimezone('Asia/Bangkok').create();
  }
  if (have.indexOf('remindSoon') === -1) {
    ScriptApp.newTrigger('remindSoon').timeBased().everyMinutes(5).create();
  }
}

/** Creates the sheets the first time the app runs, so setup() never has to be run by hand. */
function ensureReady_() {
  if (!SpreadsheetApp.getActive().getSheetByName(MEMBERS)) setup();
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('TRIGGERS_V') !== '2') { installTriggers_(); props.setProperty('TRIGGERS_V', '2'); }
}

function doGet() {
  ensureReady_();
  const hasChannel = !!(PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ID') || LINE_CHANNEL_ID);
  const props = PropertiesService.getScriptProperties();
  const linked = read_(sheet_(MEMBERS), M_COLS).filter(m => m.lineUserId).map(m => m.name);
  return ContentService.createTextOutput([
    'Minion API is ready. Sheets OK.',
    'LINE login channel: ' + (hasChannel ? 'set' : 'NOT set'),
    'Bot token: ' + (props.getProperty('LINE_ACCESS_TOKEN') ? 'set' : 'NOT set'),
    'Family group: ' + (props.getProperty('GROUP_ID') ? 'known' : 'not seen yet'),
    'Triggers: ' + ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction()).join(', '),
    'Linked LINE accounts: ' + (linked.join(', ') || 'none yet'),
    'Bookings: ' + read_(sheet_(BOOKINGS), B_COLS).length
  ].join('\n'));
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
      try { announceBooking_(clean, existing); } catch (err) { console.error(err); }
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
          { type: 'text', text: 'สวัสดีครับ 🚗 Minion มาแล้ว!\n\n' + HELP_TEXT },
          scheduleFlex_('ตารางรถ Minion', 2)
        ] });
      } else if (ev.type === 'message' && ev.message.type === 'text') {
        const text = ev.message.text.trim().toLowerCase();
        let msg = null;
        if (BOT_KEYWORDS.indexOf(text) !== -1) msg = scheduleFlex_('ตารางรถ Minion', 2);
        else if (text === 'วันนี้') msg = scheduleFlex_('คิวรถวันนี้', 1);
        else if (text === 'พรุ่งนี้') msg = scheduleFlex_('คิวรถพรุ่งนี้', 1, 1);
        else if (text === 'สัปดาห์นี้' || text === 'อาทิตย์นี้' || text === '7 วัน') msg = scheduleFlex_('คิวรถ 7 วัน', 7, 0, true);
        else if (text === 'ช่วยด้วย' || text === 'help' || text === 'คำสั่ง') msg = { type: 'text', text: HELP_TEXT };
        if (msg) lineApi_(token, 'reply', { replyToken: ev.replyToken, messages: [msg] });
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

function scheduleFlex_(title, days, offset, compact) {
  const members = read_(sheet_(MEMBERS), M_COLS);
  const bookings = read_(sheet_(BOOKINGS), B_COLS);
  const body = [];
  for (let k = 0; k < days; k++) {
    const i = k + (offset || 0);
    const iso = Utilities.formatDate(new Date(Date.now() + i * 86400000), 'Asia/Bangkok', 'yyyy-MM-dd');
    const label = (i === 0 ? 'วันนี้ · ' : i === 1 ? 'พรุ่งนี้ · ' : '') + dayLabel_(iso);
    const list = bookings.filter(b => b.date === iso).sort((a, b) => toMin_(a.start) - toMin_(b.start));
    body.push({ type: 'text', text: label, weight: 'bold', size: 'sm', color: '#1D3F73', margin: k ? (compact ? 'md' : 'lg') : 'none' });
    if (!list.length) body.push({ type: 'text', text: 'ว่างทั้งวัน ✅', size: compact ? 'xs' : 'sm', color: '#1E7A4C', margin: 'sm' });
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

// =====================================================================
// Group notifications: new/changed bookings, clashes, handover notes, 30-minute reminders.
// Each push counts against the LINE OA monthly message quota (per person in the group).
// =====================================================================
const HELP_TEXT = 'คำสั่งที่พิมพ์ได้\n' +
  '• จอง / ตาราง / คิว — คิวรถวันนี้และพรุ่งนี้\n' +
  '• วันนี้ — คิวรถวันนี้\n' +
  '• พรุ่งนี้ — คิวรถพรุ่งนี้\n' +
  '• สัปดาห์นี้ — คิวรถ 7 วัน\n' +
  '• คำสั่ง — ดูรายการนี้\n\n' +
  'Minion จะแจ้งในกลุ่มเมื่อ\n' +
  '📅 มีการจองใหม่หรือแก้เวลา\n' +
  '⚠️ มีการจองชนกัน\n' +
  '📝 มีโน้ตส่งต่อ เช่น แบตเหลือน้อย\n' +
  '⏰ อีก 30 นาทีถึงคิวใช้รถ (แท็กชื่อคนขับ)\n' +
  '☀️ สรุปคิวรถทุกเช้า 7 โมง (วันที่มีคนจอง)';

function dayLabel_(iso) {
  const p = iso.split('-').map(Number);
  const local = new Date(p[0], p[1] - 1, p[2]);
  return TH_DAYS[local.getDay()] + ' ' + p[2] + ' ' + TH_MONTHS[p[1] - 1];
}

function pushGroup_(messages) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('LINE_ACCESS_TOKEN'), groupId = props.getProperty('GROUP_ID');
  if (!token || !groupId || !messages.length) return;
  lineApi_(token, 'push', { to: groupId, messages: messages });
}

/** One message per save: what changed, any clash, and the handover note. */
function announceBooking_(b, old) {
  const members = read_(sheet_(MEMBERS), M_COLS);
  const nameOf = id => { const m = members.find(x => x.id === id); return m ? m.name : 'ไม่ระบุ'; };
  const slot = dayLabel_(b.date) + ' ' + b.start + '–' + b.end + ' น.' + (b.purpose ? ' · ' + b.purpose : '');
  const lines = [];
  const timeChanged = !old || old.date !== b.date || old.start !== b.start || old.end !== b.end || old.memberId !== b.memberId;
  if (!old) lines.push('📅 ' + nameOf(b.memberId) + ' จองรถ ' + slot);
  else if (timeChanged) lines.push('✏️ แก้การจองเป็น ' + nameOf(b.memberId) + ' ' + slot);
  if (timeChanged) {
    const clashes = read_(sheet_(BOOKINGS), B_COLS).filter(o => o.id !== b.id && o.date === b.date &&
      toMin_(o.start) < toMin_(b.end) && toMin_(b.start) < toMin_(o.end));
    clashes.forEach(o => lines.push('⚠️ ชนกับ ' + nameOf(o.memberId) + ' ' + o.start + '–' + o.end + ' น. คุยกันหน่อยนะ'));
  }
  if (b.note && (!old || old.note !== b.note)) lines.push('📝 ' + nameOf(b.memberId) + ' ฝากโน้ต: ' + b.note);
  if (lines.length) pushGroup_([{ type: 'text', text: lines.join('\n') }]);
}

/** Time trigger every 5 minutes: remind the driver ~30 minutes before their booking, with an @mention. */
function remindSoon() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('LINE_ACCESS_TOKEN') || !props.getProperty('GROUP_ID')) return;
  const now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm').split(' ');
  const today = now[0], nowMin = toMin_(now[1]);
  let sent = {};
  try { sent = JSON.parse(props.getProperty('REMINDED') || '{}'); } catch (e) {}
  Object.keys(sent).forEach(k => { if (sent[k] < today) delete sent[k]; });   // forget past days

  const members = read_(sheet_(MEMBERS), M_COLS);
  const due = read_(sheet_(BOOKINGS), B_COLS).filter(b => {
    const left = toMin_(b.start) - nowMin;
    return b.date === today && left > 5 && left <= 35 && !sent[b.id + '|' + b.date + '|' + b.start];
  });
  due.forEach(b => {
    const m = members.find(x => x.id === b.memberId);
    const rest = ' อีก ' + (toMin_(b.start) - nowMin) + ' นาทีถึงคิวใช้รถ Minion (' + b.start + '–' + b.end + ' น.' +
      (b.purpose ? ' · ' + b.purpose : '') + ') 🚗';
    const msg = (m && m.lineUserId)
      ? { type: 'textV2', text: '⏰ {driver}' + rest, substitution: { driver: { type: 'mention', mentionee: { type: 'user', userId: m.lineUserId } } } }
      : { type: 'text', text: '⏰ ' + (m ? m.name : 'คนขับ') + rest };
    pushGroup_([msg]);
    sent[b.id + '|' + b.date + '|' + b.start] = today;
  });
  props.setProperty('REMINDED', JSON.stringify(sent));
}
