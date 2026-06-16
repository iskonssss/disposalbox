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
  const bookingSheet = bookingSS.getSheetByName('Sheet1');
  const raw          = bookingSheet.getDataRange().getValues();

  if (raw.length <= 1) {
    SpreadsheetApp.getUi().alert('No bookings found in the booking sheet.');
    return;
  }

  const headers  = raw[0].map(h => String(h).toLowerCase().trim());
  const idxDate  = headers.indexOf('date');
  const idxBatch = headers.indexOf('batch');
  const idxType  = headers.indexOf('boxtype');
  const idxStat  = headers.indexOf('status');

  const bookings = raw.slice(1)
    .filter(r => r[0])                              // skip empty rows
    .map(r => ({
      date:    r[idxDate],
      batch:   r[idxBatch],
      boxType: String(r[idxType]).toLowerCase(),
      status:  String(r[idxStat]).toLowerCase(),
    }))
    .filter(b => b.status === 'confirmed')          // confirmed only
    .sort((a, b) => String(a.date).localeCompare(String(b.date)));

  // ── 2. Build values ───────────────────────────────────────────────────────
  const quantity = bookings.length;

  const pickupList = bookings.map((b, i) => {
    const d       = new Date(b.date + 'T00:00:00');
    const dateStr = d.toLocaleDateString('en-GB', {
      day:   'numeric',
      month: 'long',
      year:  'numeric',
    });
    const suffix = b.boxType === 'reinforced' ? ' - REINFORCE' : '';
    return `${i + 1}) ${dateStr} - ${b.batch} - 2pm${suffix}`;
  }).join('\n');

  // ── 3. Write to invoice ───────────────────────────────────────────────────
  const invoiceSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(INVOICE_TAB);

  invoiceSheet.getRange('H16').setValue(quantity); // quantity (top of merged cell)
  invoiceSheet.getRange('B18').setValue(pickupList);

  SpreadsheetApp.getUi().alert(
    `Done! ${quantity} confirmed bookings written to the invoice.`
  );
}
