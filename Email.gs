function sendSubmitterReceipt_(requestId) {
  const requestDetails = apiGetRequestDetails(requestId); 
  const recipient = requestDetails.submitterEmail; 
  
  if (!recipient) return; 
  
  const subject = `Confirmed: Purchase Request Submitted (${requestDetails.id})`;
  
  const htmlBody = `
    <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 24px; color: #1F2937; background-color: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 8px;">
      <div style="border-bottom: 2px solid #F3F4F6; padding-bottom: 16px; margin-bottom: 20px;">
        <span style="font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em; color: #10B981; background-color: #D1FAE5; padding: 4px 8px; border-radius: 4px;">
          Request Received
        </span>
        <h2 style="font-size: 20px; font-weight: 700; color: #111827; margin: 12px 0 4px 0;">
          Your request is in the approval queue
        </h2>
        <p style="font-size: 14px; color: #6B7280; margin: 0;">${requestDetails.title}</p>
      </div>
      
      <p style="font-size: 14px; color: #374151; margin-bottom: 20px;">
        Thank you, ${requestDetails.submitterName}. Your request has been successfully generated and sent to the appropriate team leader(s) for review.
      </p>

      <div style="background-color: #F9FAFB; border: 1px solid #F3F4F6; border-radius: 6px; padding: 16px; margin-bottom: 24px;">
        <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
          <tr><td style="padding: 6px 0; color: #6B7280; width: 35%;"><strong>Current Status:</strong></td><td style="padding: 6px 0; font-weight: bold; color: #D97706;">Pending Review</td></tr>
          <tr><td style="padding: 6px 0; color: #6B7280;"><strong>Total Cost:</strong></td><td style="padding: 6px 0; color: #111827;">$${Number(requestDetails.total).toFixed(2)}</td></tr>
        </table>
      </div>

      <div style="text-align: center; margin-top: 24px;">
        <a href="${requestDetails.pdfLink || '#'}" style="background-color: #FFFFFF; color: #4B5563; padding: 11px 20px; text-decoration: none; border-radius: 6px; font-size: 14px; font-weight: 600; display: inline-block; border: 1px solid #D1D5DB;">Download PDF Copy</a>
      </div>
    </div>
  `;

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    htmlBody: htmlBody,
    name: "Pathway Request System"
  });
}

function sendNextApprovalEmails_(requestId) {
  const request = getRequest_(requestId); 
  const allApprovals = getApprovals_(requestId);
  
  const targetEmail = request['Current Approver'];
  const currentApproval = allApprovals.find(a => a['Approver Email'] === targetEmail);

  if (!currentApproval) {
    Logger.log("Could not find a token for approver: " + targetEmail);
    return;
  }

  Logger.log("Dispatching email to: " + targetEmail + " with Token: " + currentApproval.Token);
  sendApprovalEmail_(request, currentApproval);
  Logger.log("Email successfully dispatched to " + targetEmail);
}

function sendApprovalEmail_(request, approval) {
  const base = getWebAppUrl_();
  const approveUrl = base + '?view=approve&requestId=' + encodeURIComponent(request['Request ID']) + '&token=' + encodeURIComponent(approval.Token);
  
  const html = HtmlService.createTemplateFromFile('EmailApproval');
  html.request = request;
  html.approval = approval;
  html.approveUrl = approveUrl;
  html.orgName = getSetting_('ORG_NAME', 'Pathway Church');
  
  MailApp.sendEmail({
    to: approval['Approver Email'],
    subject: `Action Required: Review ${request['Request ID']} - ${request['Request Title']}`,
    htmlBody: html.evaluate().getContent(),
    name: "Pathway Request System"
  });
}

function sendDecisionEmail_(requestId, status, comments) {
  const request = getRequest_(requestId);
  if (!request['Submitter Email']) return;

  const html = HtmlService.createTemplateFromFile('EmailDecision');
  html.request = request;
  html.status = status;
  html.comments = comments || '';
  html.orgName = getSetting_('ORG_NAME', 'Pathway Church');
  
  MailApp.sendEmail({
    to: request['Submitter Email'],
    subject: `[${requestId}] Request Resolution Update - ${status}`,
    htmlBody: html.evaluate().getContent(),
    name: "Pathway Request System"
  });
}