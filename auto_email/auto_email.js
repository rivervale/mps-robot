function autoEmail() {
  // File and folder IDs
  const folderIdReadyToSend = '1oV0J-u7AWjwxnByegeNf3JJpOBoBs7_T'; // Ready to send folder
  const folderIdPrintAndPost = '1B0jD5r8tSZZ8_1djhx2-MQa8g1fm35Wm'; // Print and post folder
  const folderIdSent = '1EFxENHZJSFoLdBlg-j-Zyu57JBpoA-k9'; // Sent folder
  const fileIdMpsData = '1oUv4buU-IFAy9wqTDmdF_7eF40p8uTU_X8u16ujVKYU'; // Data spreadsheet

  // Regular expressions and search strings
  const emailRegex = '([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+.[a-zA-Z0-9_-]+)'; // Matches standard email addresses
  const subjectLineRegex = '(?i)^appeal|^feedback|^re-?appeal|^urgent appeal|^urgent re-?appeal|^application'; // Matches subject line of letter

  // Emails
  const emailFromName = 'CHUA Kheng Wee Louis';
  const emailAddressMP = 'khengwee.chua@wp.sg';
  const emailAddressSA = 'andrelowwy@gmail.com';

  // Email signature
  const mailSignature = `<p>
    --<br>
    Chua Kheng Wee Louis<br>
    <strong>Member of Parliament<br>
    Sengkang GRC (Rivervale)<br></strong>
    --
  </p>
  `;

  // Email disclaimer
  const mailDisclaimer = `<p><em style='font-size = 9pt'>This message and any attachment are confidential and may be privileged or otherwise protected from disclosure. If you are not the intended recipient, please contact the sender and delete this message and any attachment from your system. If you are not the intended recipient you must not copy this message or attachment or disclose the contents to any other person.</em></p>
  `;

  // Email sent to agencies
  const mailAgencyBody = `<p>Dear Sir/Madam,</p>
  <p>Please find attached a letter I have written on behalf of my constituent.</p>
  <p>For your assistance please.</p>
  ` + mailSignature + mailDisclaimer;

  // Email sent to residents
  const mailResidentBody = `<p>Dear resident,</p>
  <p>Please find attached a copy of the letter of appeal I have written on your behalf which has already been sent to the relevant agency.</p>
  <p>No further action is required from you and the agency will respond to you directly.</p>
  ` + mailSignature + mailDisclaimer;


  // Pull files in the auto-email folder
  const matchingFilesSend = DriveApp.getFolderById(folderIdReadyToSend).getFilesByType('application/vnd.google-apps.document');

  // Create sent folder with today's date (if it does not already exist)
  let sentFolderDate;
  if (matchingFilesSend.hasNext()) {
    // Get today's date (yyyyMMdd)
    const formattedDate = Utilities.formatDate(new Date(), 'GMT+8', 'yyyyMMdd');
    const sentFolder = DriveApp.getFolderById(folderIdSent);
    if (sentFolder.getFoldersByName(formattedDate).hasNext()) {
      // Get the sent folder with today's date if it already exists
      sentFolderDate = sentFolder.getFoldersByName(formattedDate).next();
    } else {
      // Otherwise create it
      sentFolderDate = sentFolder.createFolder(formattedDate);
    }
  } else {
    return; // Exit the whole function if there are no files in the auto-email folder
  }

  // Open data spreadsheet
  const ss = SpreadsheetApp.openById(fileIdMpsData);
  const registrationSheet = ss.getSheetByName('Registration responses');
  const caseRange = registrationSheet.getRange(2, 1, registrationSheet.getLastRow() - 1);

  // Rolling log to be emailed at the end after execution
  let mailLog = `<h1>Virtual Kiwi Auto-Email Log</h1>
  <p>The following emails were successfully sent</p>
  <table border='1' style='border-collapse: collapse;'>
    <tr>
      <th scope="col">Case</th>
      <th scope="col">Filename</th>
      <th scope="col">Subject line</th>
      <th scope="col">Agency email(s)</th>
      <th scope="col">Resident's email(s)</th>
    </tr>
  `;

  // Counter
  let counter = 0;

  // Iterate through each file
  while (matchingFilesSend.hasNext()) {
    // Declaring variables
    let mailSubject = 'Appeal for Assistance';
    let mailAgency = '';
    let mailResident = '';
    
    // Identify elements within file
    const workingFile = matchingFilesSend.next();
    let workingDoc = DocumentApp.openById(workingFile.getId());
    let workingDocBody = workingDoc.getBody();
    const workingDocFooter = workingDoc.getFooter().getParent().getChild(4);

    // Identify and log case number and file name
    const fileName = workingFile.getName();
    const caseNumber = fileName.slice(0, fileName.search(/-/g));
    console.log('Case: ' + caseNumber); // Logs case number
    console.log('File: ' + fileName); // Logs file name
    mailLog += `<tr>
    <td>${caseNumber}</td>
    <td>${fileName}</td>
    `;

    // Search for letter subject
    const mailSubjectRangeElement = workingDocBody.findText(subjectLineRegex);
    if (mailSubjectRangeElement) {
      mailSubject = toTitleCase(mailSubjectRangeElement.getElement().getText());
      console.log('Subject: ' + mailSubject); // Logs subject line
      mailLog += `<td>${mailSubject}</td>
      `; // mailLogs subject line
    }

    // Search for agency email(s) in header
    const agencyEmailSearchRangeElements = workingDoc.newRange().addElementsBetween(workingDocBody.getChild(1), mailSubjectRangeElement.getElement()).build().getRangeElements(); // Create RangeElements[] for search from second line to subject line
    let agencyEmailRangeElement;
    let agencyEmailsRangeElements = []; // Documenting found emails for hiding later
    let agencyEmails = []; // Documenting found emails for hiding later
    for (const rangeElement of agencyEmailSearchRangeElements) {
      agencyEmailRangeElement = rangeElement.getElement().findText(emailRegex);
      if (agencyEmailRangeElement) {
        agencyEmailsRangeElements.push(agencyEmailRangeElement);
        agencyEmails.push(agencyEmailRangeElement.getElement().getText());
        mailAgency += agencyEmailRangeElement.getElement().getText().slice(agencyEmailRangeElement.getStartOffset(), agencyEmailRangeElement.getEndOffsetInclusive() + 1) + ', ';
      }
    }
    mailAgency = mailAgency.trim(); // Trim excess whitespace

    // If agency has email(s), email the agency
    if (mailAgency) {
      MailApp.sendEmail({
        name: emailFromName,
        subject: mailSubject,
        to: mailAgency,
        htmlBody: mailAgencyBody,
        attachments: [workingDoc.getAs('application/pdf')],
      });
      console.log('Agency email(s): ' + mailAgency); // Log agency email(s)
      mailLog += `<td>${mailAgency}</td>
      `; // mailLog agency email(s)
    } else {
      console.log('No agency email(s) found, PDF copy will be prepared for printing');
      mailLog += `<td>No email found</td>
      `;
    }

    // Search for resident's email(s)
    let residentEmailRangeElement;
    let residentEmailFound;
    if (workingDocFooter.findText('Email:')) { // Search in footer first
      const residentEmailSearchRangeElementsFooter = workingDoc.newRange().addElementsBetween(workingDocFooter.findText('Email:').getElement(), workingDocFooter.getChild(workingDocFooter.getNumChildren() - 1)).build().getRangeElements(); // Create RangeElements[] for search
      for (const rangeElement of residentEmailSearchRangeElementsFooter) {
        residentEmailRangeElement = rangeElement.getElement().findText(emailRegex);
        if (residentEmailRangeElement) {
          residentEmailFound = residentEmailRangeElement.getElement().getText().slice(residentEmailRangeElement.getStartOffset(),residentEmailRangeElement.getEndOffsetInclusive() + 1);
          if (residentEmailFound === emailAddressMP) {
            continue; // Check if MP's email picked up and ignore it
          } 
          mailResident += residentEmailFound + ', ';
        }
      }
    } else if (!mailResident) { // Search in letter body after signature if not found in footer
      const residentEmailSearchRangeElementsBody = workingDoc.newRange().addElementsBetween(workingDocBody.findText('Email:', workingDocBody.findText('Member of Parliament for Sengkang GRC')).getElement(), workingDocBody.getChild(workingDocBody.getNumChildren() - 1)).build().getRangeElements(); // Create RangeElements[] for search
      for (const rangeElement of residentEmailSearchRangeElementsBody) {
        residentEmailRangeElement = rangeElement.getElement().findText(emailRegex);
        if (residentEmailRangeElement) {
          residentEmailFound = residentEmailRangeElement.getElement().getText().slice(residentEmailRangeElement.getStartOffset(),residentEmailRangeElement.getEndOffsetInclusive() + 1);
          mailResident += residentEmailFound + ', ';
        }
      }
    }
    mailResident = mailResident.trim() // Trim excess whitespace

    // Censor agency email address in PDF for resident
    let i = 0;
    for (const rangeElement of agencyEmailsRangeElements) {
      i += 1;
      rangeElement.getElement().asText().setLinkUrl(null);
      rangeElement.getElement().asText().setForegroundColor('#ffffff');
      workingDocBody.replaceText(rangeElement.getElement().getText(), '-----' + i);
    }
    workingDoc.saveAndClose();

    // If resident has emails(s) email the resident
    if (mailResident) {
      MailApp.sendEmail({
        name: emailFromName,
        subject: mailSubject,
        to: mailResident,
        htmlBody: mailResidentBody,
        attachments: [workingDoc.getAs('application/pdf')],
      });
      console.log('Resident\'s email(s): ' + mailResident); // Log resident's email
      mailLog += `<td>${mailResident}</td>
      `; // mailLog resident's email
    } else {
      console.log('No resident email found, PDF copy will be prepared for printing');
      mailLog += `<td>No email found</td>
      </tr>
      `;
    }
    
    // If either or both emails are missing, create a PDF in 'Print and post' folder
    if (!mailAgency || !mailResident) {
      DriveApp.getFolderById(folderIdPrintAndPost).createFile(
        workingDoc.getAs('application/pdf')
      );
    }

    // Move original file to subfolder within 'Sent' folder with today's date (YYYYMMDD)
    moveFiles(workingFile.getId(), sentFolderDate.getId());

    // Restore agency email address
    workingDoc = DocumentApp.openById(workingFile.getId());
    workingDocBody = workingDoc.getBody();
    workingDocBody.editAsText().setForegroundColor('#000000');
    i = 0;
    for (const email of agencyEmails) {
      i += 1;
      workingDocBody.replaceText('-----' + i, email);
    }

    // Log in spreadsheet that case has been processed
    const foundCaseNumber = caseRange.createTextFinder(caseNumber).findNext();
    if (foundCaseNumber) {
      caseRow = foundCaseNumber.getRow();      
      const cell = registrationSheet.getRange(caseRow, 13);
      cell.setValue(cell.getValue() + 1);
    }
    
    // Wrapping up one file
    console.log('-----'); // Log seperator between cases
    counter += 1;
  }

  // Send a log to the MP
  mailLog += '</table>'; // Closes off the mailLog HTML
  MailApp.sendEmail({
    name: 'Virtual Kiwi',
    subject: 'Auto email script successfully executed',
    to: emailAddressMP,
    htmlBody: mailLog,
  });

  // Send a log to the SA, attach any files that need to be printed
  const matchingFilesPrint = DriveApp.getFolderById(folderIdPrintAndPost).getFilesByType('application/pdf');
  let matchingFilesPrintAccumulator = [];
  while (matchingFilesPrint.hasNext()) {
    matchingFilesPrintAccumulator.push(matchingFilesPrint.next().getAs('application/pdf'));
  }
  MailApp.sendEmail({
    name: 'Virtual Kiwi',
    subject: `Auto email script executed${matchingFilesPrintAccumulator.length != 0 ? `, ${matchingFilesPrintAccumulator.length} file(s) to print` : ''}`,
    to: emailAddressSA,
    htmlBody: mailLog + `${matchingFilesPrintAccumulator.length != 0 ? `<p>Attached: ${matchingFilesPrintAccumulator.length} file(s) to print</p>` : ''}`,
    attachments: matchingFilesPrintAccumulator,
  });

  // Completion log
  console.log(`Execution complete
Sent: ${counter} email(s)${matchingFilesPrintAccumulator.length != 0 ? `
To print: ${matchingFilesPrintAccumulator.length} letter(s)` : ''}`)
}

function moveFiles(sourceFileId, targetFolderId) {
  let file = DriveApp.getFileById(sourceFileId);
  let folder = DriveApp.getFolderById(targetFolderId);
  file.moveTo(folder);
}

function toTitleCase(string, ignore=['a', 'an', 'and', 'at', 'but', 'by', 'for', 'in', 'nor', 'of', 'on', 'or', 'out', 'so', 'the', 'to', 'up', 'yet'], caps=['SKTC', 'S&CC', 'NRIC', 'HDB', 'BTO', 'SBF', 'MOP', 'CPF', 'MSF', 'SSO', 'FSC', 'ICA', 'PR', 'LTVP', 'STVP', 'EP', 'FDW', 'SPF', 'TP', 'LTA', 'PMD', 'TPE', 'KPE', 'SLE', 'CTE', 'LRT', 'MRT', 'MOH', 'SKH', 'ACE', 'MOM', 'WSG', 'TADM', 'TAFEP']) {
  ignore = new Set(ignore);
  caps = new Set(caps);
  return string.replace(/\w\S*/g, (word, i) => {
    if (i && caps.has(word)) {
      return word;
    }
    word = word.toLowerCase();
    if (i && ignore.has(word)) {
      return word;
    }
    return word[0].toUpperCase() + word.slice(1);
  });
};