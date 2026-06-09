const SPREADSHEET_ID = '1S_Quxn6iEex-WlLKfui-W5Eq7TipGj58p4wg9AaGHoXNtgYXmmRpj2DhKcbTCU82HQMz6TRkUJiFVG';

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Store Management Tool')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 1. AUTHENTICATION
function login(email, password) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const data = ss.getSheetByName("Users").getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email && data[i][1].toString() === password.toString()) {
      return { success: true, role: data[i][2], email: email };
    }
  }
  return { success: false };
}

// 2. INTERFACE 1: Fetch Items & Submit Req
function getInventory() {
  const data = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Inventory").getDataRange().getValues();
  data.shift(); // remove headers
  return data;
}

function submitRequest(item, qty, user) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Requirements");
  ss.appendRow([Date.now(), user, item, qty, "Pending Principal", "Pending Store", new Date()]);
  return "Requirement Raised Successfully!";
}

// 3. INTERFACE 2 & 3: Approvals
function getPendingRequests(role) {
  const data = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Requirements").getDataRange().getValues();
  data.shift();
  if (role === "Principal") return data.filter(r => r[4] === "Pending Principal");
  if (role === "StoreManager") return data.filter(r => r[4] === "Approved" && r[5] === "Pending Store");
  return [];
}

function updateStatus(reqId, role, decision) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Requirements");
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0].toString() === reqId.toString()) {
      let col = (role === "Principal") ? 5 : 6;
      sheet.getRange(i + 1, col).setValue(decision);
      return {status: "Success", details: data[i]};
    }
  }
}

// 4. INTERFACE 4: Scrap Committee (3-Year Logic)
function getScrapData() {
  const data = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Inventory").getDataRange().getValues();
  const currentYear = new Date().getFullYear();
  data.shift();
  // Filter items where (Current Year - Procurement Year) >= 3
  return data.filter(item => (currentYear - parseInt(item[2])) >= 3);
}
