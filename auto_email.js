function autoEmail() {
  // Folder IDs
  const folderIdReadyToSend = '1oV0J-u7AWjwxnByegeNf3JJpOBoBs7_T'; // Ready to send folder
  const folderIdPrintAndCcByPost = '1B0jD5r8tSZZ8_1djhx2-MQa8g1fm35Wm'; // Print and CC by post folder
  const folderIdSent = '1EFxENHZJSFoLdBlg-j-Zyu57JBpoA-k9'; // Sent folder

  // Constants
  const emailFromName = 'CHUA Kheng Wee Louis';
  const emailAddressMP = 'khengwee.chua@wp.sg';
  const emailAddressSA = 'andrelowwy@gmail.com';
  const emailRegex = '([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+.[a-zA-Z0-9_-]+)';
  const mailAgencyBody = `<p>Dear Sir/Madam,</p>
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
  const mailResidentBody = `<p>Dear resident,</p>
  <p>Please find attached a copy of the letter of appeal I have written on your behalf which has already been sent to the relevant agency.</p>
  <p>No further action is required from you and the agency will respond to you directly.</p>
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
  let mailAgency = '';
  let mailResident = '';
  let mailSubject = '';
  let mailLog = `<h1>MPS Robot Auto-Email Log</h1>
  <p>The following emails were successfully sent</p>
  <table border='1' style='border-collapse: collapse;'>
    <tr>
      <th scope="col">File</th>
      <th scope="col">Subject line</th>
      <th scope="col">Agency's email</th>
      <th scope="col">Resident's email</th>
    </tr>
  `; // Rolling log to be emailed at the end after execution
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
    console.log('File: ' + workingFile.getName()); // Logs current file's name
    mailLog += `<tr>
    <td>${workingFile.getName()}</td>
    `; // mailLogs current file's name
    workingDoc = DocumentApp.openById(workingFile.getId());
    workingDocBody = workingDoc.getBody();
    workingDocFooter = workingDoc.getFooter().getParent().getChild(4);

    // Search for appeal subject
    bodySubjectRangeElement = workingDocBody.findText('APPEAL');
    if (bodySubjectRangeElement) {
      mailSubject = toTitleCase(bodySubjectRangeElement.getElement().getText());
      console.log('Subject: ' + mailSubject); // Logs subject line
      mailLog += `<td>${mailSubject}</td>
      `; // mailLogs subject line
    }

    // Search for agency email
    bodyEmailRangeElement = workingDocBody.findText(emailRegex);
    if (bodyEmailRangeElement) {
      mailAgency = bodyEmailRangeElement
        .getElement()
        .getText()
        .slice(
          bodyEmailRangeElement.getStartOffset(),
          bodyEmailRangeElement.getEndOffsetInclusive() + 1
        );
      console.log('To: ' + mailAgency); // Logs agency's email
      mailLog += `<td>${mailAgency}</td>
      `; // mailLogs agency's email
    }

    // Send email no. 1 to the agency
    MailApp.sendEmail({
      name: emailFromName,
      subject: mailSubject,
      to: mailAgency,
      htmlBody: mailAgencyBody,
      attachments: [workingDoc.getAs('application/pdf')],
    });

    // Search for resident's email
    footerEmailRangeElement = workingDocFooter.findText(emailRegex);
    if (footerEmailRangeElement) {
      mailResident = footerEmailRangeElement
        .getElement()
        .getText()
        .slice(
          footerEmailRangeElement.getStartOffset(),
          footerEmailRangeElement.getEndOffsetInclusive() + 1
        );
      if (mailResident == emailAddressMP) {
        mailResident = null;
      }
      console.log('Cc: ' + mailResident); // Logs resident's email
      mailLog += `<td>${mailResident}</td>
      </tr>
      `; // mailLogs resident's email
    } else {
      mailLog += `<td>No email found</td>
      </tr>
      `; // mailLogs resident's email
    }

    // Censor agency email address in PDF for resident
    if (bodyEmailRangeElement) {
      bodyEmailRangeElement.getElement().asText().setLinkUrl(null);
      workingDocBody.replaceText(mailAgency, '----------');
    }
    workingDoc.saveAndClose();

    // If resident has an email, email the resident. If resident has no email, create a PDF in '5. Print and CC by post' folder. Then move original file to 'Sent' folder.
    if (mailResident) {
      MailApp.sendEmail({ // Send email no. 2 to the resident
        name: emailFromName,
        subject: mailSubject,
        to: mailResident,
        htmlBody: mailResidentBody,
        attachments: [workingDoc.getAs('application/pdf')],
      });
    } else {
      DriveApp.getFolderById(folderIdPrintAndCcByPost).createFile( // Create PDF for printing
        workingDoc.getAs('application/pdf')
      );
    }
    moveFiles(workingFile.getId(), folderIdSent);

    // Restore agency email address
    if (bodyEmailRangeElement) {
      workingDoc = DocumentApp.openById(workingFile.getId());
      workingDocBody = workingDoc.getBody();
      workingDocBody.replaceText('----------', mailAgency);
    }
  }

  // Send a log to the MP, CC the SA
  mailLog += '</table>'; // Closes off the mailLog HTML
  MailApp.sendEmail({
    name: 'MPS Robot',
    subject: 'Auto email script successfully executed',
    to: emailAddressMP,
    cc: emailAddressSA,
    htmlBody: mailLog,
  });
  console.log('Execution log sent');

  // Consolidate all PDFs generated in '5. Print and CC by post' folder and send an alert to the SA
  let matchingFilesPrint = DriveApp.getFolderById(
    folderIdPrintAndCcByPost
  ).getFilesByType('application/pdf');
  let matchingFilesPrintAccumulator = [];
  while (matchingFilesPrint.hasNext()) {
    matchingFilesPrintAccumulator.push(
      matchingFilesPrint.next().getAs('application/pdf')
    );
  }
  if (matchingFilesPrintAccumulator.length != 0) {
    MailApp.sendEmail({
      name: 'MPS Robot',
      subject: 'Please print and CC the resident(s) by post',
      to: emailAddressSA,
      body: 'Please print and CC the resident(s) by post',
      attachments: matchingFilesPrintAccumulator,
    });
    console.log('SA alerted that hard copies need to be printed and posted');
  }
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
