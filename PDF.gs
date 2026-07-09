function createRequestPdf_(requestId) {
  const request = getRequest_(requestId);
  const items = getLineItems_(requestId);
  
  const ss = getSs_();
  
  const cleanSubmittedDate = request['Submitted At'] instanceof Date ? 
    Utilities.formatDate(request['Submitted At'], ss.getSpreadsheetTimeZone(), "MM/dd/yyyy") : "N/A";
    
  const cleanNeededDate = request['Date Needed'] instanceof Date ? 
    Utilities.formatDate(request['Date Needed'], ss.getSpreadsheetTimeZone(), "MM/dd/yyyy") : (request['Date Needed'] || "N/A");
  
  const template = HtmlService.createTemplateFromFile('PdfTemplate');
  template.request = request;
  template.items = items;
  template.orgName = getSetting_('ORG_NAME', 'Pathway Church');
  
  // PASS THE CLEAN FORMATTED DATES TO THE HTML TEMPLATE HERE:
  template.cleanSubmittedDate = cleanSubmittedDate;
  template.cleanNeededDate = cleanNeededDate;
  
  const html = template.evaluate().getContent();
  const blob = Utilities.newBlob(html, 'text/html', requestId + '.html').getAs('application/pdf');
  blob.setName(requestId + ' - ' + sanitizeFileName_(request['Request Title']) + '.pdf');
  
  const folderId = getSetting_('PDF_FOLDER_ID', '');
  let file;
  
  if (folderId) {
    file = DriveApp.getFolderById(folderId).createFile(blob);
  } else {
    const parentFolders = DriveApp.getFileById(getSs_().getId()).getParents();
    file = parentFolders.hasNext() ? parentFolders.next().createFile(blob) : DriveApp.createFile(blob);
  }
  
  appendObject_(APP.SHEETS.PDF_ARCHIVE, {
    'Timestamp': new Date(),
    'Request ID': requestId,
    'PDF File ID': file.getId(),
    'PDF Link': file.getUrl()
  });
  
  return { id: file.getId(), url: file.getUrl() };
}

function sanitizeFileName_(name) {
  return String(name || '').replace(/[\x00-\x1f\\/:*?"<>|]/g, '').slice(0, 60);
}