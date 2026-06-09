function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Store Management Tool')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 1. LOGIN SYSTEM
function login(email, password) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email && data[i][1] === password) {
      return { success: true, role: data[i][2], email: email };
    }
  }
  return { success: false };
}

// 2. GET ITEMS (For Interface 1)
function getItemsList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Items");
  return sheet.getRange(2, 2, sheet.getLastRow()-1, 1).getValues().flat();
}

// 3. INTERFACE 1: Raise Requirement
function submitRequirement(item, qty, userEmail) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Requirements");
  sheet.appendRow([
    "REQ-" + Utilities.getUuid().substring(0,8), 
    userEmail, item, qty, "Pending", "Awaiting Principal", new Date(), ""
  ]);
  return "Requirement submitted successfully!";
}

// 4. INTERFACE 2 & 3: Get Data for Approvals
function getRequestsForRole(role) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Requirements");
  const data = sheet.getDataRange().getValues();
  data.shift(); // Remove headers
  
  if (role === 'Principal') {
    return data.filter(r => r[4] === "Pending");
  } else if (role === 'StoreManager') {
    return data.filter(r => r[4] === "Approved" && r[5] === "Awaiting Principal");
  }
  return [];
}

// 5. INTERFACE 3: Final Approval & Stock Check
function finalStoreApprove(reqId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reqSheet = ss.getSheetByName("Requirements");
  const itemSheet = ss.getSheetByName("Items");
  
  // Update Requirement Status
  const reqData = reqSheet.getDataRange().getValues();
  for(let i=0; i<reqData.length; i++){
    if(reqData[i][0] == reqId){
      reqSheet.getRange(i+1, 6).setValue("Allotted");
      reqSheet.getRange(i+1, 8).setValue(new Date());
      return {msg: "Approved and Allotted!", id: reqId, item: reqData[i][2], qty: reqData[i][3]};
    }
  }
}

// 6. INTERFACE 4: Scrap Committee (Items older than 3 years)
function getScrapItems() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Items");
  const data = sheet.getDataRange().getValues();
  const currentYear = new Date().getFullYear();
  return data.filter(r => (currentYear - r[3]) >= 3);
}
