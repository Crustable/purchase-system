function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Request System')
    .addItem('Install / Repair System', 'installSystem')
    .addItem('Rebuild Dashboard Tab', 'rebuildDashboard')
    .addItem('Create PDF for Selected Row', 'menuCreatePdf')
    .addToUi();
}

function getSs_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_(name) {
  const ss = getSs_();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Missing sheet: ' + name + '. Click "Request System -> Install / Repair System" to fix structural schema errors.');
  return sh;
}

function headerMap_(sheetName) {
  const row = getSheet_(sheetName).getRange(1, 1, 1, getSheet_(sheetName).getLastColumn()).getValues()[0];
  return row.reduce((map, h, i) => {
    map[String(h).trim()] = i + 1;
    return map;
  }, {});
}

function rowToObject_(headers, row) {
  const obj = {};
  headers.forEach((h, i) => obj[String(h).trim()] = row[i]);
  return obj;
}

function appendObject_(sheetName, obj) {
  const sh = getSheet_(sheetName);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  sh.appendRow(headers.map(h => obj[String(h).trim()] ?? ''));
}

function updateRequest_(requestId, updates) {
  const sh = getSheet_(APP.SHEETS.REQUESTS);
  const map = headerMap_(APP.SHEETS.REQUESTS);
  const ids = sh.getRange(2, map['Request ID'], Math.max(sh.getLastRow() - 1, 1), 1).getValues().flat();
  const idx = ids.findIndex(id => String(id) === String(requestId));
  if (idx < 0) throw new Error('Request record matching key not found: ' + requestId);
  const rowNum = idx + 2;
  Object.keys(updates).forEach(key => {
    if (!map[key]) throw new Error('Unknown Requests column mapping: ' + key);
    sh.getRange(rowNum, map[key]).setValue(updates[key]);
  });
}

function getRequest_(requestId) {
  const sh = getSheet_(APP.SHEETS.REQUESTS);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(requestId)) return rowToObject_(headers, values[i]);
  }
  throw new Error('Request matching unique identifier not found: ' + requestId);
}

function getLineItems_(requestId) {
  const sh = getSheet_(APP.SHEETS.LINE_ITEMS);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  return values.slice(1).filter(r => String(r[0]) === String(requestId)).map(r => rowToObject_(headers, r));
}

function getTeamsConfig_() {
  const sh = getSheet_(APP.SHEETS.TEAMS);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  return values.slice(1).filter(r => r[0]).map(r => rowToObject_(headers, r));
}

function getOrCreateSheet_(name, schemaFields) {
  const ss = getSs_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(schemaFields);
    sh.getRange(1, 1, 1, schemaFields.length).setFontWeight('bold').setBackground('#1F4E79').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, schemaFields.length);
  }
  return sh;
}