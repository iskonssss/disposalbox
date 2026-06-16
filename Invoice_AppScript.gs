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
    .map((r, i) => ({
      id:       r[idxId],
      date:     r[idxDate],   // may be a Date object or string
      batch:    r[idxBatch],
      boxType:  String(r[idxType]).toLowerCase(),
      status:   String(r[idxStat]).toLowerCase(),
      time:     r[idxTime] || '14:00',
      invoiced: r[idxInvoiced],
      sheetRow: i + 2,        // 1-based row in sheet (header = row 1)
    }))
    .filter(r => r.id)
    .filter(b => b.status === 'confirmed' && !b.invoiced)
    .sort((a, b) => {
      // Sort by batch number ascending (e.g. Batch 24 → 25 → 26)
      const numA = parseInt(String(a.batch).replace(/[^0-9]/g, '')) || 0;
      const numB = parseInt(String(b.batch).replace(/[^0-9]/g, '')) || 0;
      return numA - numB;
    });

  // ── 2. Build values ───────────────────────────────────────────────────────
  const quantity = bookings.length;

  if (quantity === 0) {
    SpreadsheetApp.getUi().alert('No new confirmed bookings to invoice (all already marked as invoiced).');
    return;
  }

  const pickupList = bookings.map((b, i) => {
    // b.date may be a Date object from getValues(), use it directly
    const d       = b.date instanceof Date ? b.date : new Date(String(b.date).slice(0, 10) + 'T00:00:00');
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
      bookingSheet.getRange(b.sheetRow, idxInvoiced + 1).setValue(true);
    });
  }

  SpreadsheetApp.getUi().alert(
    `Done! ${quantity} confirmed bookings written to the invoice and marked as invoiced.`
  );
}
