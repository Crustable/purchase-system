/**
 * Global Settings & Configuration Matrix
 */
const APP = {
  NAME: 'Pathway Request System',
  PREFIX: 'PR',
  SHEETS: {
    REQUESTS: 'Requests',
    LINE_ITEMS: 'Line Items',
    APPROVALS: 'Approvals',
    TEAMS: 'Teams',
    PEOPLE: 'People',
    STATUS_LOG: 'Status Log',
    SETTINGS: 'Settings',
    PDF_ARCHIVE: 'PDF Archive',
    DASHBOARD: 'Dashboard'
  },
  STATUSES: {
    SUBMITTED: 'Submitted',
    PENDING_TEAM: 'Pending Team Approval',
    PENDING_EXEC: 'Pending Executive Review',
    NEEDS_INFO: 'Needs More Info',
    APPROVED: 'Approved',
    DENIED: 'Denied',
    ORDERED: 'Ordered',
    PARTIAL: 'Partially Received',
    RECEIVED: 'Received',
    CLOSED: 'Closed',
    CANCELED: 'Canceled'
  },
  DECISIONS: {
    APPROVED: 'Approved',
    DENIED: 'Denied',
    NEEDS_INFO: 'Needs More Info'
  }
};

/**
 * Safely fetches a specific value from the configuration metadata sheet.
 */
function getSetting_(key, fallback) {
  try {
    const sh = getSheet_(APP.SHEETS.SETTINGS);
    const values = sh.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]).trim() === key) return values[i][1] !== "" ? values[i][1] : fallback;
    }
  } catch (e) {
    console.warn("Settings sheet not yet installed or accessible. Using fallback.");
  }
  return fallback;
}

/**
 * Updates or appends a key-value structural setting configuration.
 */
function setSetting_(key, value) {
  const sh = getSheet_(APP.SHEETS.SETTINGS);
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) {
      sh.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value, 'Auto-generated setting variable']);
}

function getWebAppUrl_() {
  return ScriptApp.getService().getUrl();
}