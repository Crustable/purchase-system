/**
 * Maintenance.gs
 * Administrative and maintenance utility panel.
 * Combines environment architecture configuration and structural migrations.
 */

/**
 * Installs system architecture definitions seamlessly without corrupting active datasets.
 * Run this function from the IDE to safely scaffold or check database tab configurations.
 */
function installSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Table schema configurations defined precisely to fit your exact operational formats
  const architectureSchema = {
    "Requests": [
      "Request ID", "Submitted At", "Submitter Name", "Submitter Email", "Request Title", 
      "Team(s)", "Priority", "Date Needed", "Vendor", "Item Name", "Item Link", 
      "Quantity", "Unit Cost", "Total Cost", "Business Justification", "Technical Justification", 
      "Notes", "Status", "Current Approver", "Current Approval Step", "Last Updated", "PDF Link"
    ],
    "Line Items": [
      "Request ID", "Line #", "Item Name", "Vendor", "Item Link", "Quantity", "Unit Cost", "Line Total", "Notes"
    ],
    "Approvals": [
      "Request ID", "Approval Level", "Approval Type", "Team", "Approver Name", "Approver Email", 
      "Decision", "Decision Date", "Comments", "Token"
    ],
    "Settings": [
      "Setting", "Value"
    ],
    "Teams": [
      "Team", "Team Lead Name", "Team Lead Email", "Executive Approver Name", "Executive Approver Email", 
      "Requires Team Approval", "Requires Executive Approval"
    ],
    "People": [
      "Name", "Email", "Role", "Team", "Can View All", "Active"
    ],
    "Status Log": [
      "Timestamp", "Request ID", "Old Status", "New Status", "Changed By", "Notes"
    ],
    "PDF Archive": [
      "Timestamp", "Request ID", "PDF File ID", "PDF Link"
    ]
  };

  Object.keys(architectureSchema).forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    
    // Safely insert sheet block if missing
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    
    // Scaffolds default header structure if file sheet is currently completely empty
    if (sheet.getLastRow() === 0) {
      const headers = architectureSchema[sheetName];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Polish grid presentation attributes
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight("bold")
        .setBackground("#E5E7EB") // Professional slate clean tone
        .setFontColor("#1F2937");
      sheet.setFrozenRows(1);
    }
  });

  // Seed baseline administrative workspace states if Settings table is uninitialized
  const settingsSheet = ss.getSheetByName("Settings");
  if (settingsSheet && settingsSheet.getLastRow() === 1) {
    const defaultSettings = [
      ["SYSTEM_NAME", "Pathway Request System"],
      ["REQUEST_PREFIX", "PR"],
      ["NEXT_REQUEST_NUMBER", "1"],
      ["DEFAULT_EXECUTIVE_APPROVER_NAME", "Executive Approver"],
      ["DEFAULT_EXECUTIVE_APPROVER_EMAIL", Session.getActiveUser().getEmail()],
      ["PDF_FOLDER_ID", ""]
    ];
    settingsSheet.getRange(2, 1, defaultSettings.length, 2).setValues(defaultSettings);
  }

  // Trigger spreadsheet view structural refresh using Backend function
  rebuildDashboard();
}

/**
 * Legacy operational parse engine. Run one-time only to populate relational records.
 * Can be safely left alone or emptied once your historical entries map correctly.
 */
function runDataMigrationToRelationalModel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requestSheet = ss.getSheetByName("Requests");
  const lineItemSheet = ss.getSheetByName("Line Items");

  if (!requestSheet || !lineItemSheet) {
    throw new Error("Relational architecture tables must be initialized via Setup first.");
  }

  const rData = requestSheet.getDataRange().getValues();
  if (rData.length <= 1) return; // No targets available to parse

  const headers = rData[0];
  const idxId = headers.indexOf("Request ID");
  const idxItemName = headers.indexOf("Item Name");
  const idxVendor = headers.indexOf("Vendor");
  const idxLink = headers.indexOf("Item Link");
  const idxQty = headers.indexOf("Quantity");
  const idxCost = headers.indexOf("Unit Cost");
  const idxTotal = headers.indexOf("Total Cost");
  const idxNotes = headers.indexOf("Notes");

  const preparedLineItemRows = [];

  // Parse legacy cells structural signatures out into rows
  for (let i = 1; i < rData.length; i++) {
    const row = rData[i];
    const reqId = row[idxId];
    const itemName = row[idxItemName];

    if (!reqId || !itemName) continue;

    // Check if line item entry table already contains a mapping record for this request ID
    let trackingExists = false;
    const currentLIValues = lineItemSheet.getDataRange().getValues();
    for (let j = 1; j < currentLIValues.length; j++) {
      if (String(currentLIValues[j][0]).trim() === String(reqId).trim()) {
        trackingExists = true;
        break;
      }
    }

    // If missing from relational tracking, isolate and structure item components safely
    if (!trackingExists) {
      preparedLineItemRows.push([
        reqId,                     // Request ID
        1,                         // Default Line # item index position
        itemName,                  // Item Name
        row[idxVendor] || "",      // Vendor
        row[idxLink] || "",        // Item Link
        row[idxQty] || 1,          // Quantity
        row[idxCost] || row[idxTotal] || 0, // Unit Cost tracking
        row[idxTotal] || 0,        // Line Total aggregation
        row[idxNotes] || ""        // Notes column mapping entry
      ]);
    }
  }

  // Safely write migration cache blocks out to line items engine table
  if (preparedLineItemRows.length > 0) {
    lineItemSheet.getRange(
      lineItemSheet.getLastRow() + 1, 
      1, 
      preparedLineItemRows.length, 
      9
    ).setValues(preparedLineItemRows);
    
    Logger.log("Successfully migrated " + preparedLineItemRows.length + " data strings into relational objects.");
  } else {
    Logger.log("Zero migration conversions required. System mapping completely up to date.");
  }
}