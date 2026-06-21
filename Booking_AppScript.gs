const SHEET = 'Bookings';
const KEY_MAP = { id:'id', date:'date', batch:'batch', time:'time', boxtype:'boxType', status:'status', invoiced:'invoiced' };

// Singapore public holidays — update these each year
const SG_HOLIDAYS = new Set([
  // 2025
  '2025-01-01','2025-01-29','2025-01-30','2025-03-31','2025-04-18',
  '2025-05-01','2025-05-12','2025-06-06','2025-08-09','2025-10-20','2025-12-25',
  // 2026
  '2026-01-01','2026-01-29','2026-01-30','2026-03-21','2026-03-23',
  '2026-04-03','2026-05-01','2026-05-31','2026-06-07','2026-08-09',
  '2026-10-20','2026-12-25',
]);

// ── Sheet menu ────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sutrex')
    .addItem('Sync Calendar Now', 'syncAllToCalendar')
    .addSeparator()
    .addItem('Setup Auto-Sync (every 30 min)', 'installTriggers')
    .addToUi();
}

function installTriggers() {
  // Remove any existing sync triggers to avoid duplicates
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncAllToCalendar')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('syncAllToCalendar')
    .timeBased()
    .everyMinutes(30)
    .create();

  SpreadsheetApp.getUi().alert('Done! Calendar will auto-sync every 30 minutes.');
}

// ── Date helpers ──────────────────────────────────────────────────────────────

// Convert a Google Sheets Date object or ISO string to YYYY-MM-DD
function toDateISO(val) {
  if (val instanceof Date) {
    // Sheets dates are midnight local time; +12h guards against UTC offset flipping the day
    const d  = new Date(val.getTime() + 12 * 60 * 60 * 1000);
    const y  = d.getUTCFullYear();
    const mo = String(d.getUTCMonth() + 1).padStart(2, '0');
    const dy = String(d.getUTCDate()).padStart(2, '0');
    return `${y}-${mo}-${dy}`;
  }
  return String(val).slice(0, 10);
}

// Add N days to a YYYY-MM-DD string, returns YYYY-MM-DD
function addDays(isoDate, days) {
  const [y, m, d] = isoDate.split('-').map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d + days));
  return `${dt.getUTCFullYear()}-${String(dt.getUTCMonth()+1).padStart(2,'0')}-${String(dt.getUTCDate()).padStart(2,'0')}`;
}

function dayOfWeek(isoDate) {
  const [y, m, d] = isoDate.split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, d)).getUTCDay(); // 0=Sun … 6=Sat
}

function isWorkingDay(isoDate) {
  const dow = dayOfWeek(isoDate);
  return dow !== 0 && dow !== 6 && !SG_HOLIDAYS.has(isoDate);
}

// Returns YYYY-MM-DD: 1 working day before cremation date
function getDeadlineISO(cremationISO) {
  let iso = addDays(cremationISO, -1);
  while (!isWorkingDay(iso)) iso = addDays(iso, -1);
  return iso;
}

// JavaScript Date at a given SGT time (UTC+8), no DST in Singapore
function sgtDate(isoDate, hour, minute) {
  const [y, m, d] = isoDate.split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, d, hour - 8, minute, 0));
}

// ── Calendar helpers ──────────────────────────────────────────────────────────

// "Batch 25" → "#25", "Batch A1" → "#A1"
function shortBatch(batch) {
  const m = String(batch).match(/Batch\s*(A?\d+)/i);
  return m ? '#' + m[1] : String(batch);
}

// Normalize time values from Sheets (Date object, ISO string, or "HH:mm")
function normalizeSheetTime(val) {
  if (!val) return '14:00';
  if (val instanceof Date) {
    const formatted = Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm');
    return parseInt(formatted.split(':')[0], 10) < 12 ? '09:00' : '14:00';
  }
  const s = String(val);
  if (s.includes('T')) {
    const hSGT = (parseInt(s.slice(11, 13), 10) + 8) % 24;
    return hSGT < 12 ? '09:00' : '14:00';
  }
  return parseInt(s.split(':')[0], 10) < 12 ? '09:00' : '14:00';
}

// ── Calendar sync ─────────────────────────────────────────────────────────────

function syncAllToCalendar() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET);
  const raw = ss.getDataRange().getValues();
  if (raw.length <= 1) return;

  const headers  = raw[0].map(h => String(h).toLowerCase().trim());
  const idxDate  = headers.indexOf('date');
  const idxBatch = headers.indexOf('batch');
  const idxStat  = headers.indexOf('status');
  const idxTime  = headers.indexOf('time');
  const idxType  = headers.indexOf('boxtype');

  // Build deadline map: deadlineISO → array of ALL confirmed bookings for that deadline
  const deadlineMap = {};
  raw.slice(1).forEach(r => {
    if (!r[0]) return;
    if (String(r[idxStat]).toLowerCase() !== 'confirmed') return;

    const cremationISO = toDateISO(r[idxDate]);
    const deadlineISO  = getDeadlineISO(cremationISO);

    if (!deadlineMap[deadlineISO]) deadlineMap[deadlineISO] = [];
    deadlineMap[deadlineISO].push({
      batch:        String(r[idxBatch]),
      time:         normalizeSheetTime(r[idxTime]),
      boxType:      String(r[idxType]).toLowerCase(),
      cremationISO,
    });
  });

  // Fetch existing Cadaver Box events (1 month back → 13 months ahead)
  const cal        = CalendarApp.getDefaultCalendar();
  const rangeStart = new Date(); rangeStart.setMonth(rangeStart.getMonth() - 1);
  const rangeEnd   = new Date(); rangeEnd.setMonth(rangeEnd.getMonth() + 13);

  // Key existing events by "deadlineISO_hour" so we can match them precisely
  const existingByKey = {};
  cal.getEvents(rangeStart, rangeEnd)
    .filter(e => e.getTitle().includes('Cadaver Box'))
    .forEach(e => {
      const dateKey = toDateISO(e.getStartTime());
      const sgtHour = (e.getStartTime().getUTCHours() + 8) % 24;
      existingByKey[`${dateKey}_${sgtHour}`] = e;
    });

  const usedKeys = new Set();

  // Update in place where possible — only create/delete when truly needed
  Object.entries(deadlineMap).forEach(([deadlineISO, bookings]) => {
    // Standard first, then reinforced
    bookings.sort((a, b) => (a.boxType === 'standard' ? -1 : 1));

    const cremDt    = sgtDate(bookings[0].cremationISO, 12, 0);
    const cremLabel = Utilities.formatDate(cremDt, 'Asia/Singapore', 'EEEE, d MMMM yyyy');

    bookings.forEach((b, i) => {
      const hour  = 9 + i;
      const key   = `${deadlineISO}_${hour}`;
      const title = `⚫ ${shortBatch(b.batch)} Cadaver Box`;
      const desc  = `Cremation: ${cremLabel}`;

      usedKeys.add(key);

      if (existingByKey[key]) {
        // Event already exists — only write if something changed
        const ev = existingByKey[key];
        if (ev.getTitle()       !== title) ev.setTitle(title);
        if (ev.getDescription() !== desc)  ev.setDescription(desc);
      } else {
        // New event needed
        const startDt = sgtDate(deadlineISO, hour, 0);
        const endDt   = new Date(startDt.getTime() + 60 * 60 * 1000);
        const ev      = cal.createEvent(title, startDt, endDt);
        ev.setColor(CalendarApp.EventColor.GRAY);
        ev.setDescription(desc);
        ev.addPopupReminder((hour - 6) * 60); // 6am reminder
      }
    });
  });

  // Delete only events that no longer have a matching booking
  Object.entries(existingByKey).forEach(([key, ev]) => {
    if (!usedKeys.has(key)) ev.deleteEvent();
  });
}

// ── Web app (data API) ────────────────────────────────────────────────────────

function doGet(e) {
  const action = e.parameter.action || 'get';
  const ss     = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET);

  if (action === 'get') {
    const rows = ss.getDataRange().getValues();
    if (rows.length <= 1) return json([]);
    const headers = rows[0];
    const data = rows.slice(1)
      .filter(r => r[0])
      .map(r => {
        const obj = {};
        headers.forEach((h, i) => {
          const key = KEY_MAP[String(h).toLowerCase()] || String(h).toLowerCase();
          let val = r[i];
          if (key === 'date' && val instanceof Date) val = toDateISO(val);
          if (key === 'time' && val instanceof Date) {
            // Time cells stored as numeric fraction come back as Date objects
            val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm');
          }
          obj[key] = val;
        });
        return obj;
      });
    return json(data);
  }

  if (action === 'add') {
    const b = JSON.parse(e.parameter.data);
    ss.appendRow([b.id, b.date, b.batch, b.time, b.boxType, b.status, b.invoiced || false]);
    return json({ ok: true });
  }

  if (action === 'addMany') {
    const bookings = JSON.parse(e.parameter.data);
    bookings.forEach(b => ss.appendRow([b.id, b.date, b.batch, b.time, b.boxType, b.status, b.invoiced || false]));
    return json({ ok: true });
  }

  if (action === 'update') {
    const b    = JSON.parse(e.parameter.data);
    const rows = ss.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(b.id)) {
        ss.getRange(i + 1, 1, 1, 7).setValues([[b.id, b.date, b.batch, b.time, b.boxType, b.status, b.invoiced || false]]);
        break;
      }
    }
    return json({ ok: true });
  }

  if (action === 'delete') {
    const id   = e.parameter.id;
    const rows = ss.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        ss.deleteRow(i + 1);
        break;
      }
    }
    return json({ ok: true });
  }

  if (action === 'markInvoiced') {
    const ids  = JSON.parse(e.parameter.ids);
    const rows = ss.getDataRange().getValues();
    ids.forEach(id => {
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(id)) {
          ss.getRange(i + 1, 7).setValue(true);
          break;
        }
      }
    });
    return json({ ok: true });
  }

  return json({ error: 'unknown action' });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
