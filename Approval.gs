function isTruthy_(value) {
  const v = String(value).toLowerCase().trim();
  return value === true || v === 'true' || v === 'yes' || v === 'y' || v === '1';
}

function buildApprovalRows_(requestId, selectedTeams) {
  selectedTeams = Array.isArray(selectedTeams) ? selectedTeams : String(selectedTeams || '').split(',').map(t => t.trim()).filter(Boolean);
  const configs = getTeamsConfig_().filter(t => ('Active' in t) ? isTruthy_(t.Active) : true);
  const selected = configs.filter(t => selectedTeams.includes(t.Team));
  const executiveSeen = new Set();

  selected.forEach(team => {
    if (isTruthy_(team['Requires Team Approval'])) {
      appendApproval_(requestId, 1, 'Team', team.Team, team['Team Lead Name'], team['Team Lead Email']);
    }

    if (isTruthy_(team['Requires Executive Approval'])) {
      const email = team['Executive Approver Email'] || getSetting_('DEFAULT_EXECUTIVE_EMAIL', 'jake@example.com');
      const name = team['Executive Approver Name'] || getSetting_('DEFAULT_EXECUTIVE_APPROVER_NAME', 'Jake');
      if (!email) return;
      const key = email.toLowerCase().trim();

      if (!executiveSeen.has(key)) {
        executiveSeen.add(key);
        appendApproval_(requestId, 2, 'Executive', 'All Selected Teams', name, email);
      }
    }
  });
}

function appendApproval_(requestId, level, type, team, name, email) {
  if (!email) return;
  appendObject_(APP.SHEETS.APPROVALS, {
    'Request ID': requestId,
    'Approval Level': level,
    'Approval Type': type,
    'Team': team,
    'Approver Name': name,
    'Approver Email': email,
    'Decision': '',
    'Decision Date': '',
    'Comments': '',
    'Token': Utilities.getUuid()
  });
}

function approveRequest(requestId, token, decision, comments) {
  const sh = getSheet_(APP.SHEETS.APPROVALS);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const map = headerMap_(APP.SHEETS.APPROVALS);
  
  let targetRow = -1;
  let approval = null;

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][map['Request ID'] - 1]) === String(requestId) && String(values[i][map['Token'] - 1]) === String(token)) {
      targetRow = i + 1;
      approval = rowToObject_(headers, values[i]);
      break;
    }
  }

  if (targetRow < 0) throw new Error('Authorization credentials mismatched, expired, or invalid.');
  
  // Explicitly check for valid decision states to prevent ghost-data bugs
  if (approval.Decision === APP.DECISIONS.APPROVED || approval.Decision === APP.DECISIONS.DENIED) {
    return { ok: true, message: 'This localized decision step has already been filed.' };
  }

  sh.getRange(targetRow, map['Decision']).setValue(decision);
  sh.getRange(targetRow, map['Decision Date']).setValue(new Date());
  sh.getRange(targetRow, map['Comments']).setValue(comments || '');

  processApprovalState_(requestId, approval, decision, comments || '');
  rebuildDashboard();

  return { ok: true, message: decision + ' status mapping successfully parsed for item index ' + requestId };
}

function processApprovalState_(requestId, approval, decision, comments) {
  const request = getRequest_(requestId);
  const oldStatus = request.Status;

  // 1. Existing Denial Logic 
  if (decision === APP.DECISIONS.DENIED) {
    updateRequest_(requestId, { 'Status': APP.STATUSES.DENIED, 'Current Approval Step': 'Denied', 'Current Approver': '', 'Last Updated': new Date() });
    logStatus_(requestId, oldStatus, APP.STATUSES.DENIED, approval['Approver Email'], comments);
    sendDecisionEmail_(requestId, APP.STATUSES.DENIED, comments);
    return;
  }

  // 2. Existing Needs Info Logic 
  if (decision === APP.DECISIONS.NEEDS_INFO) {
    updateRequest_(requestId, { 'Status': APP.STATUSES.NEEDS_INFO, 'Current Approval Step': 'Needs More Info', 'Current Approver': request['Submitter Email'], 'Last Updated': new Date() });
    logStatus_(requestId, oldStatus, APP.STATUSES.NEEDS_INFO, approval['Approver Email'], comments);
    sendDecisionEmail_(requestId, APP.STATUSES.NEEDS_INFO, comments);
    return;
  }

  const approvals = getApprovals_(requestId);
  const level1 = approvals.filter(a => Number(a['Approval Level']) === 1);
  const level2 = approvals.filter(a => Number(a['Approval Level']) === 2);

  // 3. The Executive Guard Clause (Prevents looking backwards)
  if (Number(approval['Approval Level']) === 2) {
    const allLevel2Approved = level2.every(a => a.Decision === APP.DECISIONS.APPROVED);
    if (allLevel2Approved) {
      updateRequest_(requestId, { 'Status': APP.STATUSES.APPROVED, 'Current Approval Step': 'Approved', 'Current Approver': '', 'Last Updated': new Date() });
      logStatus_(requestId, oldStatus, APP.STATUSES.APPROVED, approval['Approver Email'], comments);
      sendDecisionEmail_(requestId, APP.STATUSES.APPROVED, comments);
    }
    return;
  }

  // 4. Level 1 Logic 
  const allLevel1Approved = level1.every(a => a.Decision === APP.DECISIONS.APPROVED);
  if (level1.length && !allLevel1Approved) {
    const pendingLeads = level1.filter(a => a.Decision !== APP.DECISIONS.APPROVED && a.Decision !== APP.DECISIONS.DENIED)
                           .map(a => a['Approver Email']).join(', ');
    updateRequest_(requestId, { 'Status': APP.STATUSES.PENDING_TEAM, 'Current Approval Step': 'Team Approval', 'Current Approver': pendingLeads, 'Last Updated': new Date() });
    logStatus_(requestId, oldStatus, APP.STATUSES.PENDING_TEAM, approval['Approver Email'], comments);
    return;
  }

  // 5. Escalation Logic (Moving from L1 to L2)
  const allLevel2Approved = level2.every(a => a.Decision === APP.DECISIONS.APPROVED);
  if (level2.length && !allLevel2Approved) {
    const pendingExecs = level2.filter(a => !a.Decision).map(a => a['Approver Email']).join(', ');
    updateRequest_(requestId, { 'Status': APP.STATUSES.PENDING_EXEC, 'Current Approval Step': 'Executive Review', 'Current Approver': pendingExecs, 'Last Updated': new Date() });
    logStatus_(requestId, oldStatus, APP.STATUSES.PENDING_EXEC, approval['Approver Email'], 'Team level signatures complete. Escalated to Executive window.');
    sendNextApprovalEmails_(requestId);
    return;
  }

  // 6. Final Approval (Fallback)
  updateRequest_(requestId, { 'Status': APP.STATUSES.APPROVED, 'Current Approval Step': 'Approved', 'Current Approver': '', 'Last Updated': new Date() });
  logStatus_(requestId, oldStatus, APP.STATUSES.APPROVED, approval['Approver Email'], comments);
  sendDecisionEmail_(requestId, APP.STATUSES.APPROVED, comments);
}

function getApprovals_(requestId) {
  const sh = getSheet_(APP.SHEETS.APPROVALS);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  return values.slice(1).filter(r => String(r[0]) === String(requestId)).map(r => rowToObject_(headers, r));
}

function getRequestForApproval(requestId, token) {
  const approvals = getApprovals_(requestId);
  const activeApproval = approvals.find(a => String(a.Token) === String(token));
  if (!activeApproval) throw new Error("Access token verification failed for requested records index mapping.");
  
  return {
    request: getRequest_(requestId),
    items: getLineItems_(requestId),
    approvalStep: activeApproval
  };
}