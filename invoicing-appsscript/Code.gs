/**
 * General Invoice — Apps Script
 * Called via fetch() from iskonssss.github.io/invoicing
 * Set API_TOKEN in: Project Settings → Script properties
 */

// ── CONFIG ────────────────────────────────────────────────────────────────────

const INVOICE_SS_ID      = '1Y7jetGJn_-irlPQ1_aeIUHDXT4x_PgQ5pAA3bZ97-44';
const INVOICE_TAB        = '12xxB Invoice -';
const INVOICES_FOLDER_ID = '1n_5BYetf6Si1zdZuYRD1WfyO7vWHA62P';
const TIMEZONE           = 'Asia/Singapore';

// Sheet structure
const FIRST_ITEM_ROW  = 17;   // first item block starts here
const ROWS_PER_ITEM   = 3;    // each item = 3 rows (label, description, extra)
const BASE_ITEM_COUNT = 1;    // sheet initially has 1 item block (rows 17-19)
const BASE_TERMS_ROW  = 25;   // terms cell row when 1 item exists
const BASE_PRINT_ROWS = 36;   // print range height with 1 item

// Cell addresses (top-left of each merged range)
const CELLS = {
  number:        'A6',   // A6:B6
  dateOfIssue:   'C6',   // C6:D6
  customerName:  'A8',   // A8:D8
  contactNumber: 'A9',   // A9:D9
  company:       'A10',  // A10:D10
  addressLine1:  'A11',  // A11:D11
  addressLine2:  'A12',  // A12:D12
  jobHeader:     'B16',  // B16:F16
};

// ── WEB APP ───────────────────────────────────────────────────────────────────

function doGet(e) {
  e = e || {};
  const params = e.parameter || {};
  if (!checkToken_(params.token)) return json({ error: 'Unauthorized' });

  const action = params.action || '';
  try {
    if (action === 'getNextNumber') return json({ nextNumberRaw: getNextInvoiceNumber_() });
    return json({ error: 'Unknown action' });
  } catch (err) {
    return json({ error: err.message });
  }
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

function checkToken_(token) {
  if (!token) return false;
  const stored = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
  return !!stored && token === stored;
}

// ── DOCUMENTS API ─────────────────────────────────────────────────────────────

function previewDocument_(data) {
  const { sheet, printRows, invoiceNum } = writeToSheet_(data);
  const blob = exportPdf_(sheet, 'A1:I' + printRows);
  return { pdfBase64: Utilities.base64Encode(blob.getBytes()), invoiceNum };
}

function saveDocument_(data) {
  const { sheet, printRows, invoiceNum } = writeToSheet_(data);
  const filename = buildFileName_(data, invoiceNum);
  const blob     = exportPdf_(sheet, 'A1:I' + printRows).setName(filename);
  const folder   = getMonthFolder_(parseDocDate_(data.dateOfIssue));

  const dupes = folder.getFilesByName(filename);
  while (dupes.hasNext()) dupes.next().setTrashed(true);

  const file = folder.createFile(blob);

  // Refresh cache and return next number
  let nextNum = null;
  try {
    nextNum = getNextInvoiceNumber_();
    PropertiesService.getScriptProperties().setProperty('NEXT_NUM_CACHE', String(nextNum));
  } catch (_) {}

  return { fileName: file.getName(), url: file.getUrl(), folder: folder.getName(), nextNumberRaw: nextNum };
}

// ── SHEET WRITE ───────────────────────────────────────────────────────────────

function writeToSheet_(data) {
  const sheet = SpreadsheetApp.openById(INVOICE_SS_ID).getSheetByName(INVOICE_TAB);
  if (!sheet) throw new Error('Tab not found: "' + INVOICE_TAB + '"');

  const items    = (data.items || []).filter(function (i) { return i.header || i.deliverables || i.notes || i.unitPrice || i.quantity; });
  const newCount = Math.max(items.length, 1);

  // ── Adjust item rows ──────────────────────────────────────────────────────

  const props        = PropertiesService.getScriptProperties();
  const currentCount = parseInt(props.getProperty('ITEM_ROW_COUNT') || BASE_ITEM_COUNT, 10);
  const currentLast  = FIRST_ITEM_ROW + currentCount * ROWS_PER_ITEM - 1;

  if (newCount > currentCount) {
    // Insert rows and copy format from first item block
    const rowsToAdd      = (newCount - currentCount) * ROWS_PER_ITEM;
    const templateRange  = sheet.getRange(FIRST_ITEM_ROW, 1, ROWS_PER_ITEM, sheet.getLastColumn());
    sheet.insertRowsAfter(currentLast, rowsToAdd);
    for (let i = currentCount; i < newCount; i++) {
      templateRange.copyTo(
        sheet.getRange(FIRST_ITEM_ROW + i * ROWS_PER_ITEM, 1),
        SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false
      );
    }
  } else if (newCount < currentCount) {
    const firstExtra = FIRST_ITEM_ROW + newCount * ROWS_PER_ITEM;
    sheet.deleteRows(firstExtra, (currentCount - newCount) * ROWS_PER_ITEM);
  }

  props.setProperty('ITEM_ROW_COUNT', String(newCount));

  // ── Computed layout positions ─────────────────────────────────────────────

  const termsRow = BASE_TERMS_ROW + (newCount - BASE_ITEM_COUNT) * ROWS_PER_ITEM;
  const printRows = BASE_PRINT_ROWS + (newCount - BASE_ITEM_COUNT) * ROWS_PER_ITEM;

  // ── Auto-generate invoice number if blank ────────────────────────────────

  const invoiceNum = (data.number || '').trim() || (cachedNextNumber_() + 'B');

  // ── Write header fields ───────────────────────────────────────────────────

  sheet.getRange(CELLS.number).setValue(invoiceNum);
  sheet.getRange(CELLS.dateOfIssue).setValue(parseDocDate_(data.dateOfIssue)).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(CELLS.customerName).setValue(data.customerName || '');
  sheet.getRange(CELLS.contactNumber).setValue(data.contactNumber || '');
  sheet.getRange(CELLS.company).setValue(data.company || '');
  sheet.getRange(CELLS.addressLine1).setValue(data.addressLine1 || '');
  sheet.getRange(CELLS.addressLine2).setValue(data.addressLine2 || '');
  sheet.getRange(CELLS.jobHeader).setValue(data.jobHeader || '');

  // ── Write items ───────────────────────────────────────────────────────────

  const showLabels = data.showItemLabels !== false;

  for (let i = 0; i < newCount; i++) {
    const base    = FIRST_ITEM_ROW + i * ROWS_PER_ITEM;  // label row
    const descRow = base + 1;                             // description row

    // Clear this item block content (B–H)
    sheet.getRange(base, 2, ROWS_PER_ITEM, 7).clearContent();

    const item = items[i] || {};

    const notesRow = base + 2;

    // Header row (row 1 of block) — show/hide based on toggle
    if (showLabels && item.header) {
      sheet.showRows(base);
      sheet.getRange('B' + base).setValue(item.header);
    } else if (showLabels) {
      sheet.showRows(base);
      sheet.getRange('B' + base).setValue('');
    } else {
      sheet.getRange('B' + base).setValue('');
      sheet.hideRows(base);
    }

    // Deliverables row (row 2 of block)
    if (item.deliverables) sheet.getRange('B' + descRow).setValue(item.deliverables);

    // Notes row (row 3 of block)
    if (item.notes) sheet.getRange('B' + notesRow).setValue(item.notes);

    // Unit price (G) and Quantity (H) — merged vertically across the 3 rows
    if (item.unitPrice) sheet.getRange('G' + base).setValue(Number(item.unitPrice));
    if (item.quantity)  sheet.getRange('H' + base).setValue(Number(item.quantity));
  }

  // ── Write terms ───────────────────────────────────────────────────────────

  if (data.termsText != null) sheet.getRange('A' + termsRow).setValue(data.termsText);

  SpreadsheetApp.flush();
  return { sheet, printRows, termsRow, invoiceNum };
}

// ── PDF EXPORT ────────────────────────────────────────────────────────────────

function exportPdf_(sheet, printRange) {
  const ss     = sheet.getParent();
  const params = [
    'format=pdf', 'gid=' + sheet.getSheetId(),
    'size=A4', 'portrait=true', 'fitw=true',
    'gridlines=false', 'printtitle=false', 'sheetnames=false',
    'pagenumbers=false', 'fzr=false',
    'top_margin=0.5', 'bottom_margin=0.5', 'left_margin=0.5', 'right_margin=0.5',
    'range=' + encodeURIComponent(printRange),
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

function buildFileName_(data, invoiceNum) {
  return (invoiceNum || data.number || 'Invoice') + ' Invoice - ' + (data.customerName || 'Customer').trim() + '.pdf';
}

function parseDocDate_(s) {
  if (!s) return new Date();
  const p = String(s).split('-');
  if (p.length === 3) return new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
  return new Date(s);
}

// ── NUMBER SERIES ─────────────────────────────────────────────────────────────

function cachedNextNumber_() {
  const props  = PropertiesService.getScriptProperties();
  const cached = parseInt(props.getProperty('NEXT_NUM_CACHE') || '0', 10);
  if (cached > 0) return cached;
  const num = getNextInvoiceNumber_();
  props.setProperty('NEXT_NUM_CACHE', String(num));
  return num;
}

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
  return max === 0 ? 1001 : max + 1;
}

function getMonthFolder_(date) {
  const root      = DriveApp.getFolderById(INVOICES_FOLDER_ID);
  const year      = String(date.getFullYear());
  const monthName = (date.getMonth() + 1) + ' ' + Utilities.formatDate(date, TIMEZONE, 'MMMM');
  return getOrCreateFolder_(getOrCreateFolder_(root, year), monthName);
}

function getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
