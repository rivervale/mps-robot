function autoEmail() {
  // Folder IDs
  const folderIdReadyToSend = '1oV0J-u7AWjwxnByegeNf3JJpOBoBs7_T'; // Ready to send folder
  const folderIdPrintAndCcByPost = '1B0jD5r8tSZZ8_1djhx2-MQa8g1fm35Wm'; // Print and CC by post folder
  const folderIdSent = '1EFxENHZJSFoLdBlg-j-Zyu57JBpoA-k9'; // Sent folder

  // Constants
  const emailAddressSA = 'andrelowwy@gmail.com';
  const emailRegex = '([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+.[a-zA-Z0-9_-]+)';
  const mailBody = `
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

  // Pull files in the auto-email folder
  let matchingFilesSend = DriveApp.getFolderById(
    folderIdReadyToSend
  ).getFilesByType('application/vnd.google-apps.document');

  // Iterate through each file
  while (matchingFilesSend.hasNext()) {
    // Identify elements within file
    workingFile = matchingFilesSend.next();
    console.log(workingFile.getName()); // Logs current file's name
    workingDoc = DocumentApp.openById(workingFile.getId());
    workingDocBody = workingDoc.getBody();
    workingDocFooter = workingDoc.getFooter().getParent().getChild(4);

    // Search for agency email
    bodyEmailRangeElement = workingDocBody.findText(emailRegex);
    if (bodyEmailRangeElement) {
      mailTo = bodyEmailRangeElement
        .getElement()
        .getText()
        .slice(
          bodyEmailRangeElement.getStartOffset(),
          bodyEmailRangeElement.getEndOffsetInclusive() + 1
        );
      console.log('To: ' + mailTo);

      // Censor agency email in preparation for PDF
      bodyEmailRangeElement.getElement().asText().setLinkUrl(null);
      workingDocBody.replaceText(mailTo, '----------');
    }

    // Search for resident's email
    footerEmailRangeElement = workingDocFooter.findText(emailRegex);
    if (footerEmailRangeElement) {
      mailCc = footerEmailRangeElement
        .getElement()
        .getText()
        .slice(
          footerEmailRangeElement.getStartOffset(),
          footerEmailRangeElement.getEndOffsetInclusive() + 1
        );
      if (mailCc == 'khengwee.chua@wp.sg') {
        mailCc = null;
      }
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
      attachments: [workingDoc.getAs('application/pdf')],
    });

    // Move file to 'Sent' folder and create PDF in '5. Print and CC by post' folder if resident has no email address
    if (mailCc) {
      moveFiles(workingFile.getId(), folderIdSent);
    } else {
      DriveApp.getFolderById(folderIdPrintAndCcByPost).createFile(
        workingDoc.getAs('application/pdf')
      );
      moveFiles(workingFile.getId(), folderIdSent);
    }

    // Restore agency email address
    if (bodyEmailRangeElement) {
      workingDoc = DocumentApp.openById(workingFile.getId());
      workingDocBody = workingDoc.getBody();
      workingDocBody.replaceText('----------', mailTo);
    }
  }

  // Consolidate all PDFs generated in '5. Print and CC by post' folder and send an alert to the SA
  let matchingFilesPrint =
    DriveApp.getFolderById(folderIdPrintAndCcByPost).getFilesByType(
      'application/pdf'
    );

  let matchingFilesPrintArray = [];

  while (matchingFilesPrint.hasNext()) {
    matchingFilesPrintArray.push(matchingFilesPrint.next().getAs('application/pdf'));
  }

  MailApp.sendEmail(
    emailAddressSA,
    'Please print and CC the resident(s) by post',
    'Please print and CC the resident(s) by post',
    {
      name: 'MPS Robot',
      attachments: matchingFilesPrintArray,
    }
  );
}

function moveFiles(sourceFileId, targetFolderId) {
  let file = DriveApp.getFileById(sourceFileId);
  let folder = DriveApp.getFolderById(targetFolderId);
  file.moveTo(folder);
}

function toTitleCase(str) {
  return str.replace(/\w\S*/g, function (txt) {
    return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
  });
}
