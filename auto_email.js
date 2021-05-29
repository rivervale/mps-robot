function autoEmail() {
  // Initialising variables
  let mailTo = '';
  let mailCc = '';
  let mailSubject = '';
  let workingFile;
  let workingDoc;
  let workingDocBody;
  let bodyEmailRangeElement;
  let bodySubjectRangeElement;
  let footerEmailRangeElement;
  let mailBody = `
    <p>Dear Sir/Madam,</p>
    <p>Please find attached a letter I have written on behalf of my constituent.</p>
    <p>For your assistance please.</p>
    <p>
      --<br>
      Chua Kheng Wee Louis<br>
      <strong>Member of Parliament<br>
      Sengkang GRC (Rivervale)<br></strong>
      --
    </p>
    <p><em style='font-size = 9pt'>This message and any attachment are confidential and may be privileged or otherwise protected from disclosure. If you are not the intended recipient, please contact the sender and delete this message and any attachment from your system.  If you are not the intended recipient you must not copy this message or attachment or disclose the contents to any other person.</em></p>
  `;

  // Pull files in the auto-email folder
  let matchingFiles = DriveApp.getFolderById('1oV0J-u7AWjwxnByegeNf3JJpOBoBs7_T').getFiles();

  // Iterate through each file
  while (matchingFiles.hasNext()) {

    // Identify elements within file
    workingFile = matchingFiles.next();
    console.log(workingFile.getName()); // Logs current file's name
    workingDoc = DocumentApp.openById(workingFile.getId());
    workingDocBody = workingDoc.getBody();
    workingDocFooter = workingDoc.getFooter().getParent().getChild(4);

    // Search for agency email
    bodyEmailRangeElement = workingDocBody.findText('([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)');
    if (bodyEmailRangeElement) {
      mailTo = bodyEmailRangeElement.getElement().getText().slice(bodyEmailRangeElement.getStartOffset(),bodyEmailRangeElement.getEndOffsetInclusive() + 1);
      console.log('To: ' + mailTo);

      // Censor agency email in preparation for PDF
      bodyEmailRangeElement.getElement().asText().setLinkUrl(null);
      workingDocBody.replaceText(mailTo, '----------');
    }

    // Search for resident's email
    footerEmailRangeElement = workingDocFooter.findText('([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)');
    if (footerEmailRangeElement) {
      mailCc = footerEmailRangeElement.getElement().getText().slice(footerEmailRangeElement.getStartOffset(),footerEmailRangeElement.getEndOffsetInclusive() + 1);
      console.log('Cc: ' + mailCc);
    }

    // Search for appeal subject
    bodySubjectRangeElement = workingDocBody.findText('APPEAL');
    if (bodySubjectRangeElement) {
      mailSubject = toTitleCase(bodySubjectRangeElement.getElement().getText());
      console.log('Subject: ' + mailSubject);
    }

    // Send the email
    workingDoc.saveAndClose();
    MailApp.sendEmail('', mailSubject, '', {
      name: 'CHUA Kheng Wee Louis',
      bcc: mailTo,
      cc: mailCc,
      htmlBody: mailBody,
      attachments: [workingDoc.getAs('application/pdf')]
    });

    // Move the sent file if resident has been CCed. Leave in folder if resident has not been CCed.
    if (footerEmailRangeElement) {
      moveFiles(workingFile.getId(), '1EFxENHZJSFoLdBlg-j-Zyu57JBpoA-k9');
    }

    // Restore agency email address
    if (bodyEmailRangeElement) {
      workingDoc = DocumentApp.openById(workingFile.getId());
      workingDocBody = workingDoc.getBody();
      workingDocBody.replaceText('----------', mailTo);
    }
  }
}

function moveFiles(sourceFileId, targetFolderId) {
  let file = DriveApp.getFileById(sourceFileId);
  let folder = DriveApp.getFolderById(targetFolderId);
  file.moveTo(folder);
}

function toTitleCase(str) {
  return str.replace(
    /\w\S*/g,
    function(txt) {
      return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
    }
  );
}