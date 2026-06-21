/**
 * Alex Sutrex Invoice — Apps Script
 *
 * Attached to the Alex Sutrex Invoice spreadsheet.
 * Called via ?action=invoice from the booking frontend.
 * Also callable from the sheet menu.
 */

// ── CONFIG ────────────────────────────────────────────────────────────────────

const BOOKING_SS_ID      = '1ZqZ_4pwdiGwltgQtnUBBYMngo1lS4h8AGshMyInMtoY';
const INVOICE_SS_ID      = '1FzR1mWku-Vn2fMUDThWybrWYDo9QpBfga3bKZZaQRL0';
const INVOICE_TAB        = '10xxB Invoice - Alex Sutrex Singapore';
const INVOICES_FOLDER_ID = '1n_5BYetf6Si1zdZuYRD1WfyO7vWHA62P';

// ── SHEET MENU ────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sutrex')
    .addItem('Refresh Invoice from Bookings', 'refreshInvoice')
    .addToUi();
}

// ── WEB APP ───────────────────────────────────────────────────────────────────

function doGet(e) {
  e = e || {};
  const action = (e.parameter && e.parameter.action) || '';

  if (action === 'invoice') {
    try {
      const result = runRefreshInvoice();
      return json({ ok: true, invoiceNum: result.invoiceNum, quantity: result.quantity });
    } catch (err) {
      return json({ ok: false, error: err.message });
    }
  }

  return json({ error: 'Unknown action' });
}

// ── SUTREX INVOICE ────────────────────────────────────────────────────────────

function refreshInvoice() {
  const result     = runRefreshInvoice();
  const monthLabel = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM yyyy');
  SpreadsheetApp.getUi().alert(
    `Done!\nInvoice: ${result.invoiceNum}\n${result.quantity} booking(s) invoiced.\n\nPDF saved to Drive → ${monthLabel}`
  );
}

function runRefreshInvoice() {
  const bookingSS    = SpreadsheetApp.openById(BOOKING_SS_ID);
  const bookingSheet = bookingSS.getSheetByName('Bookings');
  const raw          = bookingSheet.getDataRange().getValues();

  if (raw.length <= 1) throw new Error('No bookings found.');

  const headers     = raw[0].map(h => String(h).toLowerCase().trim());
  const idxId       = headers.indexOf('id');
  const idxDate     = headers.indexOf('date');
  const idxBatch    = headers.indexOf('batch');
  const idxType     = headers.indexOf('boxtype');
  const idxStat     = headers.indexOf('status');
  const idxTime     = headers.indexOf('time');
  const idxInvoiced = headers.indexOf('invoiced');

  const bookings = raw.slice(1)
    .map((r, i) => ({
      id:       r[idxId],
      date:     r[idxDate],
      batch:    r[idxBatch],
      boxType:  String(r[idxType]).toLowerCase(),
      status:   String(r[idxStat]).toLowerCase(),
      time:     r[idxTime] || '14:00',
      invoiced: r[idxInvoiced],
      sheetRow: i + 2,
    }))
    .filter(r => r.id)
    .filter(b => b.status === 'confirmed' && !b.invoiced)
    .sort((a, b) => {
      const numA = parseInt(String(a.batch).replace(/[^0-9]/g, '')) || 0;
      const numB = parseInt(String(b.batch).replace(/[^0-9]/g, '')) || 0;
      return numA - numB;
    });

  const quantity = bookings.length;
  if (quantity === 0) throw new Error('No uninvoiced confirmed bookings found.');

  const pickupList = bookings.map((b, i) => {
    const d       = b.date instanceof Date ? b.date : new Date(String(b.date).slice(0, 10) + 'T00:00:00');
    const dateStr = d.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });
    const timeStr = String(b.time) === '09:00' ? '9am' : '2pm';
    const suffix  = b.boxType === 'reinforced' ? ' - REINFORCE' : '';
    return `${i + 1}) ${dateStr} - ${b.batch} - ${timeStr}${suffix}`;
  }).join('\n');

  const nextInvoiceNum = `${getNextInvoiceNumber_()}B`;

  const invoiceSS    = SpreadsheetApp.openById(INVOICE_SS_ID);
  const invoiceSheet = invoiceSS.getSheetByName(INVOICE_TAB);

  invoiceSheet.getRange('A6').setValue(nextInvoiceNum);
  invoiceSheet.getRange('C6').setValue(new Date());
  invoiceSheet.getRange('H16').setValue(quantity);
  invoiceSheet.getRange('B18').setValue(pickupList);

  if (idxInvoiced >= 0) {
    bookings.forEach(b => {
      bookingSheet.getRange(b.sheetRow, idxInvoiced + 1).setValue(true);
    });
  }

  SpreadsheetApp.flush();
  saveInvoiceAsPDF_(invoiceSS, nextInvoiceNum);

  return { invoiceNum: nextInvoiceNum, quantity };
}

function saveInvoiceAsPDF_(ss, invoiceNum) {
  const sheet = ss.getSheetByName(INVOICE_TAB);
  const ssId  = ss.getId();
  const gid   = sheet.getSheetId();

  const url = `https://docs.google.com/spreadsheets/d/${ssId}/export` +
    `?format=pdf&gid=${gid}&size=A4&portrait=true&fitw=true` +
    `&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false`;

  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('PDF export failed: ' + response.getContentText());
  }

  const filename = `${invoiceNum} Invoice - Alex Sutrex Singapore.pdf`;
  const blob     = response.getBlob().setName(filename);
  const folder   = getMonthFolder_(new Date());

  const existing = folder.getFilesByName(filename);
  while (existing.hasNext()) existing.next().setTrashed(true);

  folder.createFile(blob);
}

// ── HELPERS ───────────────────────────────────────────────────────────────────

// Scans Drive folder for highest {n}B filename, returns next number
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
  const monthName = (date.getMonth() + 1) + ' ' +
                    Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMMM');
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
