/**
 * KHARABOOK — Simple Udhaar / Ledger App
 * Backend: Google Sheets (data) + Google Drive (PDF storage)
 * Frontend: Apps Script HTML Service (single web app)
 *
 * SETUP:
 * 1. Open (or create) a Google Sheet.
 * 2. Extensions > Apps Script.
 * 3. Paste this file as Code.gs, and Index.html as a separate HTML file (see Index.html).
 * 4. Run the "setup" function once (select it in the dropdown, click Run).
 *    Grant permissions when asked.
 * 5. Deploy > New deployment > Web app > Execute as "Me", Access "Anyone with the link".
 * 6. Open the deployed URL. That's your app.
 * 7. (Optional) Run "createMonthlyTrigger" once to auto-generate bills on the 1st of every month.
 */

const MAX_USERS = 500;
const DRIVE_FOLDER_NAME = 'KharaBook_Documents';

// PASTE YOUR SHEET ID HERE.
// Get it from your Sheet's URL: https://docs.google.com/spreadsheets/d/THIS_PART_HERE/edit
// This is required because getActiveSpreadsheet() returns null when the script runs
// as a deployed web app (there is no "active" sheet in that execution context).
const SHEET_ID = '1FpbOxPUqSEqfFVVdwZgFT3El9rumH_WrJQdh3bbHBWQ';

function getSS() {
  if (!SHEET_ID) {
    throw new Error("Please set SHEET_ID.");
  }
  return SpreadsheetApp.openById(SHEET_ID);
}

// Sheet schemas — used both by setup() and by the auto-healing getOrCreateSheet() below,
// so a missing sheet never causes a null-pointer error; it just gets created on first use.
const SHEET_SCHEMAS = {
  Users: ['UserID', 'Name', 'Phone', 'Balance', 'CreatedDate', 'Aadhar_URL', 'PAN_URL', 'Cheque_URL', 'Other_URL'],
  Transactions: ['TxnID', 'UserID', 'Date', 'Type', 'Amount', 'Note', 'BillPDF_URL', 'RunningBalance'],
  MonthlyBills: ['BillID', 'UserID', 'Month', 'PDF_URL', 'GeneratedDate']
};

function getOrCreateSheet(name) {
  const ss = getSS();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(SHEET_SCHEMAS[name]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ---------------- SETUP ----------------
// Optional to run manually — getOrCreateSheet() means the app builds sheets on first use
// automatically. Running this once is still a good way to confirm SHEET_ID is correct
// and Drive/Sheets permissions are granted.
function setup() {
  Object.keys(SHEET_SCHEMAS).forEach(function (name) { getOrCreateSheet(name); });
  getOrCreateRootFolder();
  safeAlert('KharaBook setup complete. Sheets and Drive folder are ready.');
}

// getUi() only works when called from a sheet menu/sidebar — not from the editor's
// Run button or from doGet/doPost. This falls back to a log entry instead of throwing.
function safeAlert(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    Logger.log(message);
  }
}

function getOrCreateRootFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function getUserFolder(userId, userName) {
  const root = getOrCreateRootFolder();
  const folderName = userId + '_' + String(userName).replace(/[^a-zA-Z0-9]/g, '_');
  const existing = root.getFoldersByName(folderName);
  if (existing.hasNext()) return existing.next();
  return root.createFolder(folderName);
}

// ---------------- WEB APP ENTRY ----------------
function doGet(e) {
  // Lets you test the backend directly by pasting a URL like this into your browser:
  //   YOUR_EXEC_URL?action=getUsers
  //   YOUR_EXEC_URL?action=getDashboardData
  // This bypasses fetch()/CORS entirely, so if this works but the app doesn't,
  // the problem is in the frontend connection, not the backend.
  if (e && e.parameter && e.parameter.action) {
    const params = e.parameter.params ? JSON.parse(e.parameter.params) : {};
    return handleApiRequest(e.parameter.action, params);
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('KharaBook')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------------- JSON API (for a separately-hosted frontend) ----------------
// The deployed web app URL accepts POST requests of the form:
//   { "action": "getUsers", "params": {} }
// and returns: { "success": true, "data": ... } or { "success": false, "error": "..." }
function doPost(e) {
  let action = '';
  try {
    const request = JSON.parse(e.postData.contents);
    action = request.action;
    return handleApiRequest(action, request.params || {});
  } catch (err) {
    return jsonOutput({ success: false, error: err.message, action: action });
  }
}

function handleApiRequest(action, p) {
  try {
    let data;
    switch (action) {
      case 'addUser':
        data = addUser(p.name, p.phone);
        break;
      case 'getUsers':
        data = getUsers();
        break;
      case 'getUserLedger':
        data = getUserLedger(p.userId);
        break;
      case 'addTransaction':
        data = addTransaction(p.userId, p.type, p.amount, p.note, p.billBase64, p.billFileName);
        break;
      case 'getDashboardData':
        data = getDashboardData();
        break;
      case 'uploadKYCDoc':
        data = uploadKYCDoc(p.userId, p.docType, p.base64Data, p.fileName);
        break;
      case 'generateMonthlyBillForUser':
        data = generateMonthlyBillForUser(p.userId, p.monthStr);
        break;
      default:
        throw new Error('Unknown action: ' + action);
    }
    return jsonOutput({ success: true, data: data });
  } catch (err) {
    return jsonOutput({ success: false, error: err.message, action: action });
  }
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------- USERS ----------------
function addUser(name, phone) {
  name = (name || '').trim();
  phone = (phone || '').trim();
  if (!name) throw new Error('Name is required.');

  const sheet = getOrCreateSheet('Users');
  const lastRow = sheet.getLastRow();
  const userCount = lastRow - 1; // minus header

  if (userCount >= MAX_USERS) {
    throw new Error('Maximum of ' + MAX_USERS + ' users reached.');
  }

  const userId = 'U' + String(userCount + 1).padStart(4, '0');
  sheet.appendRow([userId, name, phone, 0, new Date(), '', '', '', '']);
  return { userId: userId, name: name, phone: phone, balance: 0 };
}

function getUsers() {
  const sheet = getOrCreateSheet('Users');
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  return data.map(function (row) {
    const obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  });
}

function getUserById(userId) {
  const users = getUsers();
  for (let i = 0; i < users.length; i++) {
    if (users[i].UserID === userId) return users[i];
  }
  return null;
}

// ---------------- TRANSACTIONS ----------------
/**
 * type: 'CREDIT'  -> customer took goods/money on credit, balance goes UP (they owe you)
 *       'PAYMENT' -> customer paid you money, balance goes DOWN
 * billBase64: optional base64 PDF (a photo converted to PDF client-side) of the receiving bill
 */
function addTransaction(userId, type, amount, note, billBase64, billFileName) {
  amount = Number(amount);
  if (!amount || amount <= 0) throw new Error('Enter a valid amount.');
  if (type !== 'CREDIT' && type !== 'PAYMENT') throw new Error('Invalid transaction type.');

  const usersSheet = getOrCreateSheet('Users');
  const txnSheet = getOrCreateSheet('Transactions');

  const usersData = usersSheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < usersData.length; i++) {
    if (usersData[i][0] === userId) { rowIndex = i; break; }
  }
  if (rowIndex === -1) throw new Error('User not found.');

  let balance = Number(usersData[rowIndex][3]) || 0;
  balance = type === 'CREDIT' ? balance + amount : balance - amount;
  usersSheet.getRange(rowIndex + 1, 4).setValue(balance);

  let billUrl = '';
  if (billBase64) {
    const userName = usersData[rowIndex][1];
    billUrl = saveBillPDF(userId, userName, billBase64, billFileName);
  }

  const txnLastRow = txnSheet.getLastRow();
  const txnId = 'T' + String(txnLastRow).padStart(5, '0');
  txnSheet.appendRow([txnId, userId, new Date(), type, amount, note || '', billUrl, balance]);

  return { txnId: txnId, balance: balance, billUrl: billUrl };
}

function saveBillPDF(userId, userName, base64Data, fileName) {
  const folder = getUserFolder(userId, userName);
  const decoded = Utilities.base64Decode(base64Data.split(',').pop());
  const blob = Utilities.newBlob(decoded, MimeType.PDF, (fileName || 'bill') + '.pdf');
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function getUserLedger(userId) {
  const sheet = getOrCreateSheet('Transactions');
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  return data
    .filter(function (r) { return r[1] === userId; })
    .map(function (row) {
      const obj = {};
      headers.forEach(function (h, i) { obj[h] = row[i]; });
      return obj;
    })
    .reverse(); // newest first
}

// ---------------- DASHBOARD ----------------
function getDashboardData() {
  const users = getUsers();

  const totalOwed = users.reduce(function (s, u) {
    return s + (Number(u.Balance) > 0 ? Number(u.Balance) : 0);
  }, 0);

  const totalAdvance = users.reduce(function (s, u) {
    return s + (Number(u.Balance) < 0 ? -Number(u.Balance) : 0);
  }, 0);

  const topDebtors = users
    .filter(function (u) { return Number(u.Balance) > 0; })
    .sort(function (a, b) { return b.Balance - a.Balance; })
    .slice(0, 10);

  const txnSheet = getOrCreateSheet('Transactions');
  const txnData = txnSheet.getDataRange().getValues();
  const txnHeaders = txnData.shift();

  const userMap = {};
  users.forEach(function (u) { userMap[u.UserID] = u.Name; });

  const topPayments = txnData
    .filter(function (r) { return r[3] === 'PAYMENT'; })
    .map(function (row) {
      const obj = {};
      txnHeaders.forEach(function (h, i) { obj[h] = row[i]; });
      obj.Name = userMap[obj.UserID] || obj.UserID;
      return obj;
    })
    .sort(function (a, b) { return b.Amount - a.Amount; })
    .slice(0, 10);

  return {
    totalUsers: users.length,
    totalOwed: totalOwed,
    totalAdvance: totalAdvance,
    topDebtors: topDebtors,
    topPayments: topPayments
  };
}

// ---------------- KYC ----------------
function uploadKYCDoc(userId, docType, base64Data, fileName) {
  const validTypes = ['Aadhar', 'PAN', 'Cheque', 'Other'];
  if (validTypes.indexOf(docType) === -1) throw new Error('Invalid document type.');

  const usersSheet = getOrCreateSheet('Users');
  const usersData = usersSheet.getDataRange().getValues();

  let rowIndex = -1;
  for (let i = 1; i < usersData.length; i++) {
    if (usersData[i][0] === userId) { rowIndex = i; break; }
  }
  if (rowIndex === -1) throw new Error('User not found.');

  const userName = usersData[rowIndex][1];
  const folder = getUserFolder(userId, userName);
  const decoded = Utilities.base64Decode(base64Data.split(',').pop());
  const blob = Utilities.newBlob(decoded, MimeType.PDF, docType + '_' + userId + '.pdf');
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const colMap = { Aadhar: 6, PAN: 7, Cheque: 8, Other: 9 };
  usersSheet.getRange(rowIndex + 1, colMap[docType]).setValue(file.getUrl());

  return file.getUrl();
}

// ---------------- MONTHLY BILLS ----------------
function generateMonthlyBillForUser(userId, monthStr) {
  const now = new Date();
  if (!monthStr) {
    const prev = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    monthStr = Utilities.formatDate(prev, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  const parts = monthStr.split('-').map(Number);
  const year = parts[0], month = parts[1];
  const start = new Date(year, month - 1, 1);
  const end = new Date(year, month, 1);

  const user = getUserById(userId);
  if (!user) throw new Error('User not found.');

  const ledger = getUserLedger(userId).filter(function (t) {
    const d = new Date(t.Date);
    return d >= start && d < end;
  }).reverse(); // chronological in the bill

  let html = '<h2>KharaBook — Monthly Statement</h2>' +
    '<p><b>Name:</b> ' + user.Name + ' &nbsp; <b>Phone:</b> ' + user.Phone + '</p>' +
    '<p><b>Month:</b> ' + monthStr + '</p>' +
    '<table border="1" cellpadding="6" style="border-collapse:collapse;width:100%">' +
    '<tr><th>Date</th><th>Type</th><th>Amount</th><th>Note</th><th>Balance</th></tr>';

  ledger.forEach(function (t) {
    const d = Utilities.formatDate(new Date(t.Date), Session.getScriptTimeZone(), 'dd-MM-yyyy');
    html += '<tr><td>' + d + '</td><td>' + t.Type + '</td><td>' + t.Amount + '</td><td>' +
      (t.Note || '') + '</td><td>' + t.RunningBalance + '</td></tr>';
  });

  html += '</table><p style="margin-top:20px"><b>Closing Balance: Rs. ' + user.Balance + '</b></p>';

  const pdfBlob = Utilities.newBlob(html, 'text/html', 'statement.html').getAs('application/pdf');
  pdfBlob.setName(user.Name + '_' + monthStr + '_statement.pdf');

  const folder = getUserFolder(userId, user.Name);
  const file = folder.createFile(pdfBlob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const billsSheet = getOrCreateSheet('MonthlyBills');
  const billId = 'B' + new Date().getTime();
  billsSheet.appendRow([billId, userId, monthStr, file.getUrl(), new Date()]);

  return file.getUrl();
}

function generateAllMonthlyBills() {
  const users = getUsers();
  users.forEach(function (u) {
    try {
      generateMonthlyBillForUser(u.UserID);
    } catch (e) {
      Logger.log('Failed for ' + u.UserID + ': ' + e.message);
    }
  });
}

// Run once to auto-generate all bills on the 1st of every month at 6 AM
function createMonthlyTrigger() {
  ScriptApp.newTrigger('generateAllMonthlyBills')
    .timeBased()
    .onMonthDay(1)
    .atHour(6)
    .create();
  safeAlert('Monthly bill trigger created — bills will auto-generate on the 1st of each month.');
}
