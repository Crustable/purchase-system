// Add this to handle the "Preflight" browser check
function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT)
    .setHeader("Access-Control-Allow-Origin", "https://crustable.github.io")
    .setHeader("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
    .setHeader("Access-Control-Allow-Headers", "Content-Type");
}

function doGet(e) {
  const action = e.parameter.action;
  let result;

  // The Traffic Controller: route the request to the right function
  switch(action) {
    case 'getDashboard': result = apiGetDashboard(); break;
    case 'getRequests':  result = apiGetRequestsHistory(); break; // Ensure this function exists!
    case 'getDetails':   result = apiGetRecordDetails(e.parameter.id); break;
    default: result = { error: "Invalid action" };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader("Access-Control-Allow-Origin", "https://crustable.github.io");
}

function doPost(e) {
  const payload = JSON.parse(e.postData.contents);
  const result = submitRequest(payload); // Your existing function
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader("Access-Control-Allow-Origin", "https://crustable.github.io");
}

/**
 * Dynamically resolves the published Web App URL for navigation links
 */
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

// ==========================================
// 2. DATA FETCHING FOR RECORD DETAILS (SPA)
// ==========================================

/**
 * Fetches full details for a specific request ID, including its line items.
 * @param {string} requestId The ID to search for (e.g., "PR-000027")
 * @return {Object} The compiled request and items object
 */
function getRequestDetailsForView(requestId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requestsSheet = ss.getSheetByName(APP.SHEETS.REQUESTS); 
    const itemsSheet = ss.getSheetByName(APP.SHEETS.LINE_ITEMS);

    if (!requestsSheet) throw new Error("Could not find Requests sheet tab.");

    const reqData = requestsSheet.getDataRange().getValues();
    if (reqData.length < 2) return null; 

    const reqHeaders = reqData[0];
    const idIndex = reqHeaders.indexOf("Request ID"); 
    
    if (idIndex === -1) throw new Error("Could not find a column named 'Request ID'.");

    let targetRowData = null;
    for (let i = 1; i < reqData.length; i++) {
      if (String(reqData[i][idIndex]).trim() === String(requestId).trim()) {
        targetRowData = reqData[i];
        break;
      }
    }

    if (!targetRowData) return null;

    // Convert row array into a key-value object
    const requestObject = {};
    reqHeaders.forEach((header, index) => {
      let value = targetRowData[index];
      if (value instanceof Date) {
        value = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      requestObject[header] = value;
    });

    // Gather Line Items associated with this Request ID
    const associatedItems = [];
    if (itemsSheet) {
      const itemData = itemsSheet.getDataRange().getValues();
      if (itemData.length > 1) {
        const itemHeaders = itemData[0];
        const itemIdIndex = itemHeaders.indexOf("Request ID");

        if (itemIdIndex !== -1) {
          for (let j = 1; j < itemData.length; j++) {
            if (String(itemData[j][itemIdIndex]).trim() === String(requestId).trim()) {
              const itemObj = {};
              itemHeaders.forEach((header, idx) => {
                itemObj[header] = itemData[j][idx];
              });
              associatedItems.push(itemObj);
            }
          }
        }
      }
    }

    return { request: requestObject, items: associatedItems };
  } catch (error) {
    throw new Error("Server Error in getRequestDetailsForView: " + error.toString());
  }
}

// ==========================================
// 3. PHASE 1: LIGHTNING SUBMISSION LOGIC
// ==========================================

function submitRequest(payload) {
  validatePayload_(payload);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); 
  
  try {
    const requestId = nextRequestId_();
    const now = new Date();
    
    // 1. Calculate totals
    const items = payload.items || [];
    const total = items.reduce((sum, item) => sum + (Number(item.quantity || 0) * Number(item.unitCost || 0)), 0);
    
    // 2. Write the main request row
    appendObject_(APP.SHEETS.REQUESTS, {
      'Request ID': requestId,
      'Submitted At': now,
      'Submitter Name': payload.submitterName,
      'Submitter Email': payload.submitterEmail,
      'Request Title': payload.title,
      'Team(s)': (payload.teams || []).join(', '),
      'Priority': payload.priority,
      'Date Needed': payload.dateNeededBy || '',
      'Total Cost': total,
      'Business Justification': payload.businessJustification || '',
      'Technical Justification': payload.technicalJustification || '',
      'Notes': payload.notes || '',
      'Status': APP.STATUSES.SUBMITTED,
      'Current Approval Step': 'Submitted',
      'Current Approver': '',
      'PDF Link': '', // Left blank for Phase 2 Background task
      'Last Updated': now
    });
    
    // 3. BATCH WRITE LINE ITEMS (Lightning Fast)
    if (items.length > 0) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const itemsSheet = ss.getSheetByName(APP.SHEETS.LINE_ITEMS);
      const itemRows = items.map((item, index) => [
        requestId,
        index + 1,
        item.itemName,
        item.vendor || '',
        item.itemLink || '',
        Number(item.quantity || 0),
        Number(item.unitCost || 0),
        Number(item.quantity || 0) * Number(item.unitCost || 0),
        item.notes || ''
      ]);
      const startRow = itemsSheet.getLastRow() + 1;
      itemsSheet.getRange(startRow, 1, itemRows.length, itemRows[0].length).setValues(itemRows);
    }
    
    // 4. Setup initial routing and update status instantly
    buildApprovalRows_(requestId, payload.teams || []);
    const approvals = getApprovals_(requestId);
    const level1 = approvals.filter(a => Number(a['Approval Level']) === 1);
    
    const targetStatus = level1.length > 0 ? APP.STATUSES.PENDING_TEAM : APP.STATUSES.PENDING_EXEC;
    const targetStep = level1.length > 0 ? 'Waiting for team leader signature' : 'Waiting for director approval';
    const currentApproverEmails = level1.length > 0 
      ? level1.map(a => a['Approver Email']).join(', ')
      : approvals.filter(a => Number(a['Approval Level']) === 2).map(a => a['Approver Email']).join(', ');
    
    updateRequest_(requestId, {
      'Status': targetStatus, 
      'Current Approval Step': targetStep,
      'Current Approver': currentApproverEmails,
      'Last Updated': now
    });
    
    logStatus_(requestId, APP.STATUSES.SUBMITTED, targetStatus, 'system', 'Request moved into review status.');
    
    // Optional: Only triggers if rebuild function exists
    try { rebuildDashboard(); } catch(e) {}
    
    // 5. INSTANT RETURN (User sees success in 1-2 seconds!)
    return { ok: true, requestId: requestId };
    
  } catch (err) {
    Logger.log("CRITICAL ERROR: " + err.message);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 4. PHASE 2: BACKGROUND PROCESSOR (TRIGGERED)
// ==========================================

/**
 * Runs automatically every minute to generate PDFs and send emails
 * without making the end-user wait on the loading screen.
 */
function processBackgroundTasks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requestsSheet = ss.getSheetByName(APP.SHEETS.REQUESTS);
  
  const data = requestsSheet.getDataRange().getValues();
  if(data.length < 2) return;
  const headers = data[0];
  
  const idIndex = headers.indexOf('Request ID');
  const pdfIndex = headers.indexOf('PDF Link');
  const statusIndex = headers.indexOf('Status');
  
  if (idIndex === -1 || pdfIndex === -1) return;
  
  // Look for any row that has been submitted but is missing a PDF link
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const reqId = row[idIndex];
    const pdfLink = row[pdfIndex];
    const status = row[statusIndex];
    
    // If a request exists, has moved past 'Submitted', but has no PDF yet:
    if (reqId && pdfLink === "" && status !== APP.STATUSES.SUBMITTED) {
      Logger.log("Processing background tasks for: " + reqId);
      
      try {
        // 1. Generate the heavy PDF
        const pdf = createRequestPdf_(reqId);
        
        // 2. Write the PDF link to the sheet (Column index + 1 for 1-based grid)
        requestsSheet.getRange(i + 1, pdfIndex + 1).setValue(pdf.url);
        
        // 3. Send the heavy emails
        sendSubmitterReceipt_(reqId);
        sendNextApprovalEmails_(reqId);
        
        Logger.log("Successfully completed background processing for: " + reqId);
        
      } catch (e) {
        Logger.log("Error processing background tasks for " + reqId + ": " + e.message);
      }
    }
  }
}

// ==========================================
// 5. UTILITY FUNCTIONS
// ==========================================

function validatePayload_(payload) {
  if (!payload) throw new Error('The request form cannot be completely empty.');
  if (!payload.submitterName) throw new Error('Please fill out this required field: Submitter Name');
  if (!payload.submitterEmail) throw new Error('Please fill out this required field: Submitter Email');
  if (!payload.title) throw new Error('Please fill out this required field: Request Title');
  if (!payload.priority) throw new Error('Please fill out this required field: Priority');
  if (!payload.teams || !payload.teams.length) throw new Error('Please select at least one department.');
  if (!payload.items || !payload.items.length) throw new Error('You must add at least one item to your request.');
}

function nextRequestId_() {
  const next = Number(getSetting_('NEXT_REQUEST_NUMBER', '1'));
  setSetting_('NEXT_REQUEST_NUMBER', String(next + 1));
  return APP.PREFIX + '-' + Utilities.formatString('%06d', next);
}

function logStatus_(requestId, previousStatus, newStatus, actorEmail, comment) {
  appendObject_(APP.SHEETS.STATUS_LOG, {
    'Timestamp': new Date(),
    'Request ID': requestId,
    'Old Status': previousStatus,
    'New Status': newStatus,
    'Changed By': actorEmail,
    'Notes': comment || ''
  });
}

function menuCreatePdf() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sh.getName() !== APP.SHEETS.REQUESTS) throw new Error('Open and select an index mapping inside the Requests sheet row first.');
  const row = sh.getActiveRange().getRow();
  if (row < 2) throw new Error('Cannot run action mapping calculations across a label header row context.');
  const requestId = sh.getRange(row, 1).getValue();
  const pdf = createRequestPdf_(requestId);
  updateRequest_(requestId, {'PDF Link': pdf.url, 'Last Updated': new Date()});
  SpreadsheetApp.getUi().alert('PDF Link created/refreshed: ' + pdf.url);
}
