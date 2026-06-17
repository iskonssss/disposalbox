// ID of your BOOKING spreadsheet (not the invoice one)
const BOOKING_SS_ID = '1ZqZ_4pwdiGwltgQtnUBBYMngo1lS4h8AGshMyInMtoY';

// Tab name in the INVOICE spreadsheet where the invoice lives
const INVOICE_TAB = '10xxB Invoice - Alex Sutrex Singapore';

// Root folder ID of your invoices in Google Drive
const INVOICES_FOLDER_ID = '1n_5BYetf6Si1zdZuYRD1WfyO7vWHA62P';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sutrex')
    .addItem('Refresh Invoice from Bookings', 'refreshInvoice')
    .addToUi();
}

// ── Invoice number helper ─────────────────────────────────────────────────────

function getLatestInvoiceNumber() {
  const root = DriveApp.getFolderById(INVOICES_FOLDER_ID);
  let maxNum = 1000;

  const yearFolders = root.getFolders();
  while (yearFolders.hasNext()) {
    const yearFolder = yearFolders.next();
    if (yearFolder.getName() === 'Template') continue;

    const monthFolders = yearFolder.getFolders();
    while (monthFolders.hasNext()) {
      const monthFolder = monthFolders.next();

      const files = monthFolder.getFiles();
      while (files.hasNext()) {
        const name  = files.next().getName();
        const match = name.match(/^(\d+)B/i);
        if (match) {
          const num = parseInt(match[1], 10);
          if (num > maxNum) maxNum = num;
        }
      }
    }
  }

  return maxNum;
}

// ── Folder helpers ────────────────────────────────────────────────────────────

function getOrCreateFolder(parent, name) {
  const iter = parent.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : parent.createFolder(name);
}

function getMonthFolder() {
  const now        = new Date();
  const year       = String(now.getFullYear());
  const monthNum   = now.getMonth() + 1; // 1–12
  const monthName  = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMMM'); // e.g. "June"
  const monthLabel = `${monthNum} ${monthName}`; // e.g. "6 June"

  const root       = DriveApp.getFolderById(INVOICES_FOLDER_ID);
  const yearFolder = getOrCreateFolder(root, year);
  return getOrCreateFolder(yearFolder, monthLabel);
}

// ── PDF export ────────────────────────────────────────────────────────────────

function saveInvoiceAsPDF(invoiceNum) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(INVOICE_TAB);
  const ssId  = ss.getId();
  const gid   = sheet.getSheetId();

  const url = `https://docs.google.com/spreadsheets/d/${ssId}/export` +
    `?format=pdf` +
    `&gid=${gid}` +
    `&size=A4` +
    `&portrait=true` +
    `&fitw=true` +
    `&sheetnames=false` +
    `&printtitle=false` +
    `&pagenumbers=false` +
    `&gridlines=false` +
    `&fzr=false`;

  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('PDF export failed: ' + response.getContentText());
  }

  const filename = `${invoiceNum} Invoice - Alex Sutrex Singapore.pdf`;
  const blob     = response.getBlob().setName(filename);
  const folder   = getMonthFolder();

  // Replace existing file with same name (avoid Drive duplicates)
  const existing = folder.getFilesByName(filename);
  while (existing.hasNext()) existing.next().setTrashed(true);

  folder.createFile(blob);
}

// ── Main refresh function ─────────────────────────────────────────────────────

function refreshInvoice() {
  // ── 1. Read bookings ──────────────────────────────────────────────────────
  const bookingSS    = SpreadsheetApp.openById(BOOKING_SS_ID);
  const bookingSheet = bookingSS.getSheetByName('Bookings');
  const raw          = bookingSheet.getDataRange().getValues();

  if (raw.length <= 1) {
    SpreadsheetApp.getUi().alert('No bookings found in the booking sheet.');
    return;
  }

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

  // ── 2. Build pickup list ──────────────────────────────────────────────────
  const quantity = bookings.length;

  if (quantity === 0) {
    SpreadsheetApp.getUi().alert('No new confirmed bookings to invoice (all already marked as invoiced).');
    return;
  }

  const pickupList = bookings.map((b, i) => {
    const d       = b.date instanceof Date ? b.date : new Date(String(b.date).slice(0, 10) + 'T00:00:00');
    const dateStr = d.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });
    const timeStr = (String(b.time) === '09:00') ? '9am' : '2pm';
    const suffix  = b.boxType === 'reinforced' ? ' - REINFORCE' : '';
    return `${i + 1}) ${dateStr} - ${b.batch} - ${timeStr}${suffix}`;
  }).join('\n');

  // ── 3. Get next invoice number from Drive ─────────────────────────────────
  const latestNum      = getLatestInvoiceNumber();
  const nextInvoiceNum = `${latestNum + 1}B`;

  // ── 4. Write to invoice sheet ─────────────────────────────────────────────
  const invoiceSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(INVOICE_TAB);

  invoiceSheet.getRange('A6').setValue(nextInvoiceNum);
  invoiceSheet.getRange('C6').setValue(new Date());
  invoiceSheet.getRange('H16').setValue(quantity);
  invoiceSheet.getRange('B18').setValue(pickupList);

  // ── 5. Mark bookings as invoiced in booking sheet ─────────────────────────
  if (idxInvoiced >= 0) {
    bookings.forEach(b => {
      bookingSheet.getRange(b.sheetRow, idxInvoiced + 1).setValue(true);
    });
  }

  // ── 6. Flush writes then export PDF ──────────────────────────────────────
  SpreadsheetApp.flush();
  const monthLabel = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM yyyy');
  saveInvoiceAsPDF(nextInvoiceNum);

  SpreadsheetApp.getUi().alert(
    `Done!\nInvoice: ${nextInvoiceNum}\n${quantity} booking(s) invoiced.\n\nPDF saved to Drive → ${monthLabel}`
  );
}
