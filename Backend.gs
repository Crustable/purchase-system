/**
 * Backend.gs
 * Core application runtime engine.
 * Combines front-end data pipelines and spreadsheet dashboard rendering.
 */

function apiGetDashboard() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
  } catch (e) {
    throw new Error("System busy. Please try refreshing again.");
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requestSheet = ss.getSheetByName("Requests");
    
    const stats = { total: 0, submitted: 0, pendingTeam: 0, pendingExec: 0, approved: 0, completed: 0, totalValue: 0 };
    const recentRequests = [];

    if (!requestSheet) return { stats, recentRequests };

    const data = requestSheet.getDataRange().getValues();
    if (data.length <= 1) return { stats, recentRequests };

    const headers = data[0];
    const idxId = headers.indexOf("Request ID");
    const idxTitle = headers.indexOf("Request Title");
    const idxStatus = headers.indexOf("Status");
    const idxTotal = headers.indexOf("Total Cost");
    const idxSubmitter = headers.indexOf("Submitter Name");
    const idxUpdated = headers.indexOf("Last Updated");

    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const status = idxStatus !== -1 ? String(row[idxStatus]).trim() : "";
      const cost = idxTotal !== -1 ? parseFloat(row[idxTotal]) || 0 : 0;

      stats.total++;
      stats.totalValue += cost;

      if (status === "Submitted") stats.submitted++;
      else if (status === "Pending Team Approval") stats.pendingTeam++;
      else if (status === "Pending Executive Approval") stats.pendingExec++;
      else if (status === "Approved") stats.approved++;
      else if (status === "Completed") stats.completed++;

      if (recentRequests.length < 15) {
        let formattedDate = "";
        if (idxUpdated !== -1 && row[idxUpdated]) {
          const rawDate = new Date(row[idxUpdated]);
          if (!isNaN(rawDate.getTime())) {
            formattedDate = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "MM/dd/yyyy");
          }
        }

        recentRequests.push({
          id: idxId !== -1 ? row[idxId] : "",
          title: idxTitle !== -1 ? row[idxTitle] : "",
          submitter: idxSubmitter !== -1 ? row[idxSubmitter] : "",
          status: status,
          total: cost.toFixed(2),
          lastUpdated: formattedDate
        });
      }
    }
    return { stats, recentRequests };
  } finally {
    lock.releaseLock();
  }
}

function apiGetRequestDetails(requestId) {
  if (!requestId) throw new Error("Invalid Request ID reference.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rSheet = ss.getSheetByName("Requests");
  if (!rSheet) throw new Error("Requests database tab is missing.");
  const rData = rSheet.getDataRange().getValues();
  const rHeaders = rData[0];
  
  const idxId = rHeaders.indexOf("Request ID");
  const idxTitle = rHeaders.indexOf("Request Title");
  const idxName = rHeaders.indexOf("Submitter Name");
  const idxEmail = rHeaders.indexOf("Submitter Email");
  const idxTeams = rHeaders.indexOf("Team(s)");
  const idxPriority = rHeaders.indexOf("Priority");
  const idxDate = rHeaders.indexOf("Date Needed");
  const idxTotal = rHeaders.indexOf("Total Cost");
  const idxBus = rHeaders.indexOf("Business Justification");
  const idxTech = rHeaders.indexOf("Technical Justification");
  const idxStatus = rHeaders.indexOf("Status");
  const idxPdf = rHeaders.indexOf("PDF Link");

  let row = null;
  for (let i = 1; i < rData.length; i++) {
    if (String(rData[i][idxId]).trim() === requestId.trim()) {
      row = rData[i];
      break;
    }
  }

  if (!row) throw new Error("Request record " + requestId + " not found.");

  const liSheet = ss.getSheetByName("Line Items");
  const liData = liSheet ? liSheet.getDataRange().getValues() : [];
  const parsedLineItems = [];
  
  if (liData.length > 1) {
    const liHeaders = liData[0];
    const liIdIdx = liHeaders.indexOf("Request ID");
    const liNameIdx = liHeaders.indexOf("Item Name");
    const liVendorIdx = liHeaders.indexOf("Vendor");
    const liLinkIdx = liHeaders.indexOf("Item Link");
    const liQtyIdx = liHeaders.indexOf("Quantity");
    const liCostIdx = liHeaders.indexOf("Unit Cost");

    for (let i = 1; i < liData.length; i++) {
      if (String(liData[i][liIdIdx]).trim() === requestId.trim()) {
        parsedLineItems.push({
          name: liData[i][liNameIdx] || "",
          vendor: liData[i][liVendorIdx] || "N/A",
          link: liData[i][liLinkIdx] || "",
          quantity: parseFloat(liData[i][liQtyIdx]) || 0,
          unitCost: parseFloat(liData[i][liCostIdx]) || 0
        });
      }
    }
  }

  const rawDateNeeded = row[idxDate];
  let cleanNeededString = "Not Specified";
  if (rawDateNeeded && !isNaN(new Date(rawDateNeeded).getTime())) {
    cleanNeededString = Utilities.formatDate(new Date(rawDateNeeded), ss.getSpreadsheetTimeZone(), "MM/dd/yyyy");
  }

  return {
    id: row[idxId],
    title: row[idxTitle],
    submitterName: row[idxName],
    submitterEmail: row[idxEmail],
    total: row[idxTotal],
    departments: row[idxTeams],
    priority: row[idxPriority],
    dateNeededBy: cleanNeededString,
    businessJustification: row[idxBus],
    technicalJustification: row[idxTech] || "",
    status: row[idxStatus],
    pdfLink: row[idxPdf] || null,
    items: parsedLineItems
  };
}

function executeApprovalDecision(requestId, status, comments) {
  const now = new Date();
  let operatorEmail = 'system.gatekeeper';
  try { operatorEmail = Session.getActiveUser().getEmail(); } catch(e) {}
  
  let userFriendlyStepName = '';
  if (status === 'Approved') userFriendlyStepName = 'Approved / Cleared for Order';
  else if (status === 'Denied') userFriendlyStepName = 'Rejected / Order Terminated';
  else if (status === 'More Info') userFriendlyStepName = 'Returned for Clarification';

  updateRequest_(requestId, {
    'Status': status,
    'Current Approval Step': userFriendlyStepName,
    'Current Approver': status === 'More Info' ? 'Original Submitter' : '',
    'Last Updated': now
  });
  
  logStatus_(requestId, 'Pending Review', status, operatorEmail, comments);
  rebuildDashboard();
  return { success: true };
}

function rebuildDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dashSheet = ss.getSheetByName("Dashboard");
  
  if (!dashSheet) {
    dashSheet = ss.insertSheet("Dashboard");
  } else {
    dashSheet.clear();
    dashSheet.setHiddenGridlines(false);
  }

  const uiData = apiGetDashboard();
  const stats = uiData.stats;

  const matrixData = [
    ["Metric", "Value", "Notes"],
    ["Total Requests", stats.total, "All time tracking records"],
    ["Total Requested", stats.totalValue, "Total cost of all requested items"],
    ["Submitted", stats.submitted, "Waiting for initial tracking assignment"],
    ["Pending Team Approval", stats.pendingTeam, "Waiting for team leader signature"],
    ["Pending Executive Approval", stats.pendingExec, "Waiting for director approval"],
    ["Approved", stats.approved, "Fully validated across steps"],
    ["Completed", stats.completed, "Order completed and placed"]
  ];

  const destinationRange = dashSheet.getRange(1, 1, matrixData.length, 3);
  destinationRange.setValues(matrixData);

  dashSheet.setColumnWidth(1, 220);
  dashSheet.setColumnWidth(2, 140);
  dashSheet.setColumnWidth(3, 280);

  dashSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#1F2937").setFontColor("#FFFFFF").setHorizontalAlignment("center");
  dashSheet.getRange(2, 2, matrixData.length - 1, 1).setFontWeight("bold").setHorizontalAlignment("right");
  dashSheet.getRange("B3").setNumberFormat("$#,##0.00");
  dashSheet.setFrozenRows(1);
}