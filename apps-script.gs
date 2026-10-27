const APP_SHEET = 'Payment Tracker Data';

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(APP_SHEET);
  const stored = sheet ? sheet.getRange('B1').getValue() : '';
  let data = null;
  try { data = stored ? JSON.parse(stored) : null; } catch (_) {}
  return jsonResponse({ ok: true, data: data });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action !== 'save' || !body.data) throw new Error('Invalid save request');
    saveData(body.data);
    return jsonResponse({ ok: true });
  } catch (error) {
    return jsonResponse({ ok: false, error: error.message });
  }
}

function saveData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    saveState(ss, data);
    saveCustomers(ss, data.c || []);
  } finally {
    lock.releaseLock();
  }
}

function saveState(ss, data) {
  let sheet = ss.getSheetByName(APP_SHEET) || ss.insertSheet(APP_SHEET);
  sheet.getRange('A1').setValue('App data (JSON)');
  sheet.getRange('B1').setValue(JSON.stringify(data));

  const settings = data.s || {};
  sheet.getRange('A3:D3').setValues([['Setting', 'Value', 'Setting', 'Value']]);
  sheet.getRange('A4:D10').setValues([
    ['Factory', settings.factory || '', 'Country code', settings.cc || ''],
    ['Reminder gap (days)', settings.gap || 0, 'Minimum due', settings.min || 0],
    ['UPI ID', settings.upi || '', 'Payee', settings.payee || ''],
    ['Account holder', settings.acn || '', 'Account number', settings.acc || ''],
    ['IFSC', settings.ifsc || '', 'Bank / branch', settings.bk || ''],
    ['Add payment details', settings.addPay ? 'Yes' : 'No', 'Include UPI link', settings.link ? 'Yes' : 'No'],
    ['Last backed up', new Date(), '', '']
  ]);
}

function saveCustomers(ss, customers) {
  let sheet = ss.getSheetByName('Customers') || ss.insertSheet('Customers');
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 9).setValues([[
    'Name', 'Phone', 'Place', 'Amount Due', 'Due Since', 'Language', 'Note', 'Reminders Sent', 'Last Reminder'
  ]]);
  if (!customers.length) return;
  const rows = customers.map(function(customer) {
    return [
      customer.name || '', customer.phone || '', customer.place || '', Number(customer.due || 0),
      customer.since || '', customer.lang || '', customer.note || '', Number(customer.n || 0),
      customer.last || ''
    ];
  });
  sheet.getRange(2, 1, rows.length, 9).setValues(rows);
  sheet.getRange(1, 1, 1, 9).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 9);
}

function jsonResponse(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
