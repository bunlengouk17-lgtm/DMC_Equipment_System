/**
 * DMC AV Booking Schedule -> Google Sheet (shared data source)
 *
 * This replaces the earlier write-only version. It now supports:
 *   - doGet   : returns every ACTIVE booking as JSON, so every visitor's
 *               browser can load the same real, shared list of reservations.
 *   - doPost  : action "create" appends a new booking (default if no
 *               action is given, for backwards compatibility).
 *               action "cancel" marks an existing booking as cancelled
 *               instead of deleting the row, so you keep a full history.
 *
 * REDEPLOY STEPS (you're updating an existing deployment, not making a new one):
 * 1. Open the "Booking Schedule" spreadsheet -> Extensions -> Apps Script.
 * 2. Select all the existing code and replace it with everything below.
 * 3. Save (Ctrl/Cmd+S).
 * 4. Click Deploy -> Manage deployments.
 * 5. Click the pencil/edit icon on your existing deployment.
 * 6. Under "Version," choose "New version," then click Deploy.
 *    (Do NOT create a brand new deployment - that would give you a new
 *    URL and you'd have to update both webpages again. Editing the
 *    existing deployment keeps the same URL you already gave the site.)
 * 7. That's it - the same URL now supports read + write.
 *
 * If your sheet already has rows from the old version, this script will
 * automatically add the two new columns (Items JSON, Status) the next
 * time it runs, and treats any existing rows as "Active."
 */

const SHEET_NAME = 'Sheet1';
const HEADERS = [
  'Timestamp',
  'Booking ID',
  'Name',
  'Course',
  'Class / Year',
  'Equipment',
  'Check-out Date',
  'Check-in Date',
  'Item Count',
  'Items JSON',
  'Status'
];

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME)
    || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}

function ensureHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    return;
  }
  // Backfill any missing columns for sheets created by the older version.
  const existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (existing.length < HEADERS.length) {
    sheet.getRange(1, existing.length + 1, 1, HEADERS.length - existing.length)
      .setValues([HEADERS.slice(existing.length)]);
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const sheet = getSheet();
    ensureHeaders(sheet);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return jsonResponse({ status: 'success', bookings: [] });

    const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
    const bookings = [];

    values.forEach(row => {
      const [timestamp, bookingId, name, course, className, equipment, checkOut, checkIn, itemCount, itemsJson, status] = row;
      if (!bookingId) return;
      if (status === 'Cancelled') return;

      let items = [];
      try { items = itemsJson ? JSON.parse(itemsJson) : []; } catch (err) { items = []; }

      bookings.push({
        id: String(bookingId),
        name: String(name || ''),
        course: String(course || ''),
        className: String(className || ''),
        checkOut: formatDateCell(checkOut),
        checkIn: formatDateCell(checkIn),
        items: items,
        createdAt: timestamp instanceof Date ? timestamp.toISOString() : String(timestamp || '')
      });
    });

    return jsonResponse({ status: 'success', bookings: bookings });
  } catch (err) {
    return jsonResponse({ status: 'error', message: String(err) });
  }
}

function formatDateCell(val) {
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  return String(val || '');
}

function doPost(e) {
  try {
    const sheet = getSheet();
    ensureHeaders(sheet);
    const data = JSON.parse(e.postData.contents);
    const action = data.action || 'create';

    if (action === 'cancel') {
      return cancelBooking(sheet, data.bookingId);
    }
    return createBooking(sheet, data);
  } catch (err) {
    return jsonResponse({ status: 'error', message: String(err) });
  }
}

function createBooking(sheet, data) {
  sheet.appendRow([
    new Date(),
    data.bookingId || '',
    data.name || '',
    data.course || '',
    data.className || '',
    data.equipment || '',
    data.checkOut || '',
    data.checkIn || '',
    data.itemCount || '',
    data.itemsJson || '[]',
    'Active'
  ]);
  return jsonResponse({ status: 'success' });
}

function cancelBooking(sheet, bookingId) {
  if (!bookingId) return jsonResponse({ status: 'error', message: 'Missing bookingId' });
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonResponse({ status: 'error', message: 'No bookings found' });

  const idColumn = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  for (let i = 0; i < idColumn.length; i++) {
    if (String(idColumn[i][0]) === String(bookingId)) {
      const rowNum = i + 2;
      sheet.getRange(rowNum, HEADERS.length).setValue('Cancelled');
      return jsonResponse({ status: 'success' });
    }
  }
  return jsonResponse({ status: 'error', message: 'Booking not found' });
}
