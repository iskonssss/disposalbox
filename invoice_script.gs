// ID of your BOOKING spreadsheet (not the invoice one)
const BOOKING_SS_ID = '1ZqZ_4pwdiGwltgQtnUBBYMngo1lS4h8AGshMyInMtoY';

// Tab name in the INVOICE spreadsheet where the invoice lives
const INVOICE_TAB = '10xxB Invoice - Alex Sutrex Singapore';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sutrex')
    .addItem('Refresh Invoice from Bookings', 'refreshInvoice')
    .addToUi();
}

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
    .filter(r => r[0])
    .map(r => ({
      id:       r[idxId],
      date:     r[idxDate],
      batch:    r[idxBatch],
      boxType:  String(r[idxType]).toLowerCase(),
      status:   String(r[idxStat]).toLowerCase(),
      time:     r[idxTime] || '14:00',
      invoiced: r[idxInvoiced],
      rowIndex: raw.indexOf(r),   // 0-based index into raw (includes header row)
    }))
    .filter(b => b.status === 'confirmed' && !b.invoiced)
    .sort((a, b) => String(a.date).localeCompare(String(b.date)));

  // ── 2. Build values ───────────────────────────────────────────────────────
  const quantity = bookings.length;

  if (quantity === 0) {
    SpreadsheetApp.getUi().alert('No new confirmed bookings to invoice (all already marked as invoiced).');
    return;
  }

  const pickupList = bookings.map((b, i) => {
    const d       = new Date(b.date + 'T00:00:00');
    const dateStr = d.toLocaleDateString('en-GB', {
      day:   'numeric',
      month: 'long',
      year:  'numeric',
    });
    const timeStr = (String(b.time) === '09:00') ? '9am' : '2pm';
    const suffix  = b.boxType === 'reinforced' ? ' - REINFORCE' : '';
    return `${i + 1}) ${dateStr} - ${b.batch} - ${timeStr}${suffix}`;
  }).join('\n');

  // ── 3. Write to invoice ───────────────────────────────────────────────────
  const invoiceSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(INVOICE_TAB);

  invoiceSheet.getRange('H16').setValue(quantity);
  invoiceSheet.getRange('B18').setValue(pickupList);

  // ── 4. Mark included bookings as invoiced in booking sheet ────────────────
  if (idxInvoiced >= 0) {
    bookings.forEach(b => {
      // rowIndex is 0-based index in raw array (which includes header at 0)
      // so rowIndex is already the sheet row - 1; sheet row = rowIndex + 1
      bookingSheet.getRange(b.rowIndex + 1, idxInvoiced + 1).setValue(true);
    });
  }

  SpreadsheetApp.getUi().alert(
    `Done! ${quantity} confirmed bookings written to the invoice and marked as invoiced.`
  );
}
