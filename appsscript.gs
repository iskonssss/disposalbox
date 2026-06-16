const SHEET = 'Sheet1';

// Normalise header names regardless of how user typed them in the sheet
const KEY_MAP = { id:'id', date:'date', batch:'batch', time:'time', boxtype:'boxType', status:'status' };

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
          obj[key] = r[i];
        });
        return obj;
      });
    return json(data);
  }

  if (action === 'add') {
    const b = JSON.parse(e.parameter.data);
    ss.appendRow([b.id, b.date, b.batch, b.time, b.boxType, b.status]);
    return json({ ok: true });
  }

  if (action === 'addMany') {
    const bookings = JSON.parse(e.parameter.data);
    bookings.forEach(b => ss.appendRow([b.id, b.date, b.batch, b.time, b.boxType, b.status]));
    return json({ ok: true });
  }

  if (action === 'update') {
    const b    = JSON.parse(e.parameter.data);
    const rows = ss.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(b.id)) {
        ss.getRange(i + 1, 1, 1, 6).setValues([[b.id, b.date, b.batch, b.time, b.boxType, b.status]]);
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

  return json({ error: 'unknown action' });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
