const APP_SHEET = 'Payment Tracker Data';

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(APP_SHEET);
  const stored = sheet ? sheet.getRange('B1').getValue() : '';
  let data = null;
  try { data = stored ? JSON.parse(stored) : null; } catch (_) {}
  if (!data) data = { c: [], s: {} };
  data.c = readCustomers(ss, data.c || []);
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
    data.c = readCustomers(ss, data.c || []);
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
  sheet.getRange(1, 1, 1, 11).setValues([[
    'ID', 'Name', 'Phone', 'Place', 'Amount Due', 'Due Since', 'Language', 'Note', 'Reminders Sent', 'Last Reminder', 'Updated At'
  ]]);
  if (!customers.length) {
    sheet.getRange(1, 1, 1, 11).setFontWeight('bold');
    sheet.setFrozenRows(1);
    return;
  }
  const now = new Date().toISOString();
  const rows = customers.map(function(customer) {
    return [
      customer.id || Utilities.getUuid(), customer.name || '', customer.phone || '', customer.place || '', Number(customer.due || 0),
      customer.since || '', customer.lang || '', customer.note || '', Number(customer.n || 0),
      customer.last || '', customer.updatedAt || now
    ];
  });
  sheet.getRange(2, 1, rows.length, 11).setValues(rows);
  sheet.getRange(1, 1, 1, 11).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 11);
}

function readCustomers(ss, savedCustomers) {
  const sheet = ss.getSheetByName('Customers');
  if (!sheet || sheet.getLastRow() < 2) return savedCustomers;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
  const sheetCustomers = values.map(function(row, index) {
    const name = String(row[1] || '').trim();
    if (!name) return null;
    const id = String(row[0] || '').trim() || Utilities.getUuid();
    return {
      id: id,
      name: name,
      phone: String(row[2] || '').replace(/^['=]/, ''),
      place: String(row[3] || ''),
      due: Number(row[4] || 0),
      since: formatSheetDate(row[5]),
      lang: String(row[6] || ''),
      note: String(row[7] || ''),
      n: Number(row[8] || 0),
      last: formatSheetDate(row[9]),
      updatedAt: String(row[10] || '') || new Date(Date.now() - index).toISOString()
    };
  }).filter(Boolean);

  const byId = {};
  let newestSheetEdit = 0;
  sheetCustomers.forEach(function(customer) {
    const id = String(customer.id || '');
    byId[id] = customer;
    newestSheetEdit = Math.max(newestSheetEdit, new Date(customer.updatedAt || 0).getTime());
  });
  savedCustomers.forEach(function(saved) {
    const id = String(saved.id || '').trim();
    if (!id) return;
    const sheetCustomer = byId[id];
    if (!sheetCustomer) {
      if (new Date(saved.updatedAt || 0).getTime() > newestSheetEdit) byId[id] = saved;
      return;
    }
    if (new Date(saved.updatedAt || 0).getTime() > new Date(sheetCustomer.updatedAt || 0).getTime()) {
      byId[id] = saved;
    }
  });
  return Object.keys(byId).map(function(id) { return byId[id]; });
}

function formatSheetDate(value) {
  if (!value) return '';
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const text = String(value).trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(text) ? text : '';
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== 'Customers' || e.range.getColumn() === 11) return;
  const row = e.range.getRow();
  if (row < 2 || !sheet.getRange(row, 2).getValue()) return;
  sheet.getRange(row, 11).setValue(new Date().toISOString());
}

function jsonResponse(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
