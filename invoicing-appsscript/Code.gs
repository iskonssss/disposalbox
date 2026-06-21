/**
 * General Invoice — Apps Script
 *
 * Attached to the General Invoice spreadsheet.
 * Called via fetch() from iskonssss.github.io/invoicing
 *
 * Set API_TOKEN in: Apps Script → Project Settings → Script properties
 */

// ── CONFIG ────────────────────────────────────────────────────────────────────

const INVOICE_SS_ID      = '1Y7jetGJn_-irlPQ1_aeIUHDXT4x_PgQ5pAA3bZ97-44';
const INVOICES_FOLDER_ID = '1n_5BYetf6Si1zdZuYRD1WfyO7vWHA62P';
const TIMEZONE           = 'Asia/Singapore';

const DOC_CONFIG = {
  FALLBACK_NUMBER: 1001,
  CELLS: {
    number:        'A6',
    dateOfIssue:   'C6',
    customerName:  'A8',
    contactNumber: 'A9',
    company:       'A10',
    addressLine1:  'A11',
    addressLine2:  'A12',
    jobHeader:     'B16',
    terms:         'A25',
  },
  ITEM_ROWS:     [17, 18, 19],
  ITEM_DESC_COL: 'B',
  PRINT_RANGE:   'A1:I36',
  MODES: {
    invoice: {
      label:        'Invoice',
      tabName:      '12xxB Invoice -',
      word:         'Invoice',
      numberPrefix: '',
      numberSuffix: 'B',
      partyLabel:   'Billed to',
      termsLabel:   'Terms & conditions',
      terms: {
        'Full payment':
          '- Full payment to be made before project commencement\n' +
          '- Additional costs may be incurred if changes are requested after design freeze',
        '70% upfront':
          '- 70% upfront to be made before project commencement\n' +
          '- Additional costs may be incurred if changes are requested after design freeze',
      },
      itemFields: [
        { key: 'unitPrice', label: 'Unit price (S$)', col: 'G', type: 'number' },
        { key: 'quantity',  label: 'Quantity',        col: 'H', type: 'number' },
      ],
    },
    // Quotation, Receipt, Delivery Order — add later once tab names confirmed
  },
};

// ── WEB APP ───────────────────────────────────────────────────────────────────

function doGet(e) {
  e = e || {};
  const params = e.parameter || {};

  if (!checkToken_(params.token)) return json({ error: 'Unauthorized' });

  if ((params.action || '') === 'getInitialData') {
    try { return json(getInitialData_()); }
    catch (err) { return json({ error: err.message }); }
  }

  return json({ error: 'Unknown action' });
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (_) {}

  if (!checkToken_(body.token)) return json({ error: 'Unauthorized' });

  const action = body.action || '';
  const data   = body.data   || {};

  try {
    if (action === 'previewDocument') return json(previewDocument_(data));
    if (action === 'saveDocument')    return json(saveDocument_(data));
    return json({ error: 'Unknown action' });
  } catch (err) {
    return json({ error: err.message });
  }
}

// ── TOKEN CHECK ───────────────────────────────────────────────────────────────

function checkToken_(token) {
  if (!token) return false;
  const stored = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
  return !!stored && token === stored;
}

// ── DOCUMENTS API ─────────────────────────────────────────────────────────────

function getInitialData_() {
  let nextNumberRaw = DOC_CONFIG.FALLBACK_NUMBER;
  let warning = '';
  try {
    nextNumberRaw = getNextInvoiceNumber_();
  } catch (err) {
    warning = 'Could not auto-detect the next number (' + err.message + '). Enter it manually.';
  }

  const modes = {};
  Object.keys(DOC_CONFIG.MODES).forEach(function (k) {
    const m = DOC_CONFIG.MODES[k];
    modes[k] = {
      label:        m.label,
      partyLabel:   m.partyLabel,
      termsLabel:   m.termsLabel,
      numberPrefix: m.numberPrefix,
      numberSuffix: m.numberSuffix,
      terms:        m.terms,
      itemFields:   m.itemFields.map(function (f) {
        return { key: f.key, label: f.label, type: f.type };
      }),
    };
  });

  return {
    nextNumberRaw: nextNumberRaw,
    today:         Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd'),
    modes:         modes,
    modeOrder:     Object.keys(DOC_CONFIG.MODES),
    warning:       warning,
  };
}

function previewDocument_(data) {
  const sheet = writeToSheet_(data);
  const blob  = exportPdf_(sheet);
  return { filename: buildFileName_(data), pdfBase64: Utilities.base64Encode(blob.getBytes()) };
}

function saveDocument_(data) {
  const sheet  = writeToSheet_(data);
  const blob   = exportPdf_(sheet).setName(buildFileName_(data));
  const folder = getMonthFolder_(parseDocDate_(data.dateOfIssue));

  const dupes = folder.getFilesByName(blob.getName());
  while (dupes.hasNext()) dupes.next().setTrashed(true);

  const file = folder.createFile(blob);
  return { fileName: file.getName(), url: file.getUrl(), folder: folder.getName() };
}

// ── SHEET WRITE ───────────────────────────────────────────────────────────────

function writeToSheet_(data) {
  const m     = DOC_CONFIG.MODES[data.mode] || DOC_CONFIG.MODES.invoice;
  const sheet = SpreadsheetApp.openById(INVOICE_SS_ID).getSheetByName(m.tabName);
  if (!sheet) throw new Error('Sheet tab not found: "' + m.tabName + '"');

  const C = DOC_CONFIG.CELLS;
  sheet.getRange(C.number).setValue(data.number || '');
  sheet.getRange(C.dateOfIssue).setValue(parseDocDate_(data.dateOfIssue)).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(C.customerName).setValue(data.customerName || '');
  sheet.getRange(C.contactNumber).setValue(data.contactNumber || '');
  sheet.getRange(C.company).setValue(data.company || '');
  sheet.getRange(C.addressLine1).setValue(data.addressLine1 || '');
  sheet.getRange(C.addressLine2).setValue(data.addressLine2 || '');
  sheet.getRange(C.jobHeader).setValue(data.jobHeader || '');

  // Clear item rows then fill
  const lastCol = lastItemCol_(m);
  const r0 = DOC_CONFIG.ITEM_ROWS[0];
  const r1 = DOC_CONFIG.ITEM_ROWS[DOC_CONFIG.ITEM_ROWS.length - 1];
  sheet.getRange(DOC_CONFIG.ITEM_DESC_COL + r0 + ':' + lastCol + r1).clearContent();

  (data.items || []).forEach(function (item, i) {
    if (i >= DOC_CONFIG.ITEM_ROWS.length) return;
    const row = DOC_CONFIG.ITEM_ROWS[i];
    if (item.description) sheet.getRange(DOC_CONFIG.ITEM_DESC_COL + row).setValue(item.description);
    m.itemFields.forEach(function (f) {
      const v = item[f.key];
      if (v === '' || v == null) return;
      sheet.getRange(f.col + row).setValue(f.type === 'number' ? Number(v) : v);
    });
  });

  if (data.termsText != null) sheet.getRange(C.terms).setValue(data.termsText);
  SpreadsheetApp.flush();
  return sheet;
}

function lastItemCol_(m) {
  let col = 'F';
  m.itemFields.forEach(function (f) { if (f.col > col) col = f.col; });
  return col;
}

// ── PDF EXPORT ────────────────────────────────────────────────────────────────

function exportPdf_(sheet) {
  const ss     = sheet.getParent();
  const params = [
    'format=pdf', 'gid=' + sheet.getSheetId(),
    'size=A4', 'portrait=true', 'fitw=true',
    'gridlines=false', 'printtitle=false', 'sheetnames=false',
    'pagenumbers=false', 'fzr=false',
    'top_margin=0.5', 'bottom_margin=0.5', 'left_margin=0.5', 'right_margin=0.5',
    'range=' + encodeURIComponent(DOC_CONFIG.PRINT_RANGE),
  ].join('&');

  const url  = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?' + params;
  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() !== 200) {
    throw new Error('PDF export failed (' + resp.getResponseCode() + '): ' + resp.getContentText().slice(0, 200));
  }
  return resp.getBlob();
}

function buildFileName_(data) {
  const m      = DOC_CONFIG.MODES[data.mode] || DOC_CONFIG.MODES.invoice;
  const client = (data.customerName || 'Customer').trim();
  return data.number + ' ' + m.word + ' - ' + client + '.pdf';
}

function parseDocDate_(s) {
  if (!s) return new Date();
  const p = String(s).split('-');
  if (p.length === 3) return new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
  return new Date(s);
}

// ── NUMBER SERIES ─────────────────────────────────────────────────────────────

// Shared series with Sutrex — scans same Drive folder for highest {n}B filename
function getNextInvoiceNumber_() {
  const root = DriveApp.getFolderById(INVOICES_FOLDER_ID);
  const re   = /^(\d+)\s*B\b/i;
  let max    = 0;

  const years = root.getFolders();
  while (years.hasNext()) {
    const months = years.next().getFolders();
    while (months.hasNext()) {
      const files = months.next().getFiles();
      while (files.hasNext()) {
        const m = files.next().getName().match(re);
        if (m) { const n = parseInt(m[1], 10); if (n > max) max = n; }
      }
    }
  }
  return max === 0 ? DOC_CONFIG.FALLBACK_NUMBER : max + 1;
}

function getMonthFolder_(date) {
  const root      = DriveApp.getFolderById(INVOICES_FOLDER_ID);
  const year      = String(date.getFullYear());
  const monthName = (date.getMonth() + 1) + ' ' +
                    Utilities.formatDate(date, TIMEZONE, 'MMMM');
  return getOrCreateFolder_(getOrCreateFolder_(root, year), monthName);
}

function getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
