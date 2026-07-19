/**
 * DAILY HISAB — Google Sheets backend
 * =====================================================================
 * SETUP:
 * 1. Create a new Google Sheet (this will hold one tab per salesman).
 * 2. Extensions > Apps Script. Delete any starter code, paste this whole
 *    file in, and save (name the project anything, e.g. "Hisab Backend").
 * 3. Click Deploy > New deployment > gear icon > "Web app".
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    Click Deploy, authorize when prompted, then copy the "Web app URL".
 * 4. Paste that URL into APPSSCRIPT_URL in hisab-sheets.html.
 * 5. Ignore/delete the default "Sheet1" tab — it's not treated as a
 *    salesman. Every OTHER tab name = one salesman.
 * =====================================================================
 */

const HEADERS = ['Date','MS_Initial','MS_Final','MS_Price','MS_Test',
  'HSD_Initial','HSD_Final','HSD_Price','HSD_Test',
  'UPI_Initial','UPI_Final',
  'Cash_500','Cash_200','Cash_100','Cash_50','Cash_20','Cash_10','Cash_Coins',
  'Credit','DTPlus','Other','UpdatedAt'];

function doGet(e){
  const action = e.parameter.action;
  if (action === 'salesmen') return jsonOut(getSalesmenList());
  if (action === 'entries') return jsonOut(getEntries(e.parameter.salesman));
  return jsonOut({error:'unknown action'});
}

function doPost(e){
  const body = JSON.parse(e.postData.contents);
  if (body.action === 'addSalesman') return jsonOut(addSalesman(body.name));
  if (body.action === 'saveEntry') return jsonOut(saveEntry(body.salesman, body.entry));
  if (body.action === 'deleteEntry') return jsonOut(deleteEntry(body.salesman, body.date));
  return jsonOut({error:'unknown action'});
}

function jsonOut(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name, createIfMissing){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet && createIfMissing){
    sheet = ss.insertSheet(name);
    sheet.getRange('A:A').setNumberFormat('@'); // keep dates as plain text, never auto-converted
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getSalesmenList(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets().map(s => s.getName()).filter(n => n !== 'Sheet1');
}

function addSalesman(name){
  name = (name || '').trim();
  if (!name) return {ok:false, error:'empty name'};
  getSheet(name, true);
  return {ok:true};
}

function formatDate(v){
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v);
}

function rowToEntry(r){
  return {
    date: formatDate(r[0]),
    products: {
      MS:  {initial:r[1]||0, final:r[2]||0, price:r[3]||0, test:r[4]||0},
      HSD: {initial:r[5]||0, final:r[6]||0, price:r[7]||0, test:r[8]||0}
    },
    upi: {initial:r[9]||0, final:r[10]||0},
    cash: {n500:r[11]||0, n200:r[12]||0, n100:r[13]||0, n50:r[14]||0, n20:r[15]||0, n10:r[16]||0, coins:r[17]||0},
    credit: r[18]||0,
    dtplus: r[19]||0,
    other: r[20]||0
  };
}

function getEntries(salesman){
  const sheet = getSheet(salesman, false);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  return data.slice(1)
    .filter(r => r[0])
    .map(rowToEntry)
    .sort((a,b) => b.date.localeCompare(a.date))
    .slice(0, 180);
}

function saveEntry(salesman, entry){
  const sheet = getSheet(salesman, true);
  sheet.getRange('A:A').setNumberFormat('@'); // belt-and-braces: keep column A as text
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i=1; i<data.length; i++){
    if (formatDate(data[i][0]) === entry.date){ rowIndex = i+1; break; }
  }
  const row = [
    entry.date,
    entry.products.MS.initial, entry.products.MS.final, entry.products.MS.price, entry.products.MS.test,
    entry.products.HSD.initial, entry.products.HSD.final, entry.products.HSD.price, entry.products.HSD.test,
    entry.upi.initial, entry.upi.final,
    entry.cash.n500, entry.cash.n200, entry.cash.n100, entry.cash.n50, entry.cash.n20, entry.cash.n10, entry.cash.coins,
    entry.credit, entry.dtplus, entry.other,
    new Date()
  ];
  if (rowIndex > 0) sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  else sheet.appendRow(row);
  return {ok:true};
}

function deleteEntry(salesman, date){
  const sheet = getSheet(salesman, false);
  if (!sheet) return {ok:false};
  const data = sheet.getDataRange().getValues();
  for (let i=1; i<data.length; i++){
    if (formatDate(data[i][0]) === date){ sheet.deleteRow(i+1); break; }
  }
  return {ok:true};
}
