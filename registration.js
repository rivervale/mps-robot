function onFormSubmit(e) {
  // Inspired by: https://howtogapps.com/google-form-script-to-autofill-and-email-a-doc-template/

  // Case sheet template handling
  let templateDoc = DriveApp.getFileById('10taSkDqvqppcGIOPfvYWTpdBrROQmRsgkoA5eRDKwGc');
  let newTempFile = templateDoc.makeCopy(); //create a copy
  let openDoc = DocumentApp.openById(newTempFile.getId()); //open the new template document for editing
  let body = openDoc.getBody();
  let firstFooter = openDoc.getFooter().getParent().getChild(4);

  // Data spreadsheet handling
  let ss = SpreadsheetApp.openById('1oUv4buU-IFAy9wqTDmdF_7eF40p8uTU_X8u16ujVKYU');
  const registrationSheet = ss.getSheetByName('Registration responses');
  
  // Get the responses triggered by onFormSubmit
  let items = e.response.getItemResponses();

  // Assign all form responses to variables
  let queueNumber = items[0].getResponse().padStart(2, '0'); //pads queue number to 2 digits
  let name = items[1].getResponse();
  let nric = items[2].getResponse()[0] + '####' + items[2].getResponse().slice(5);
  let dateOfBirth = items[3].getResponse();
  let gender = items[4].getResponse();
  let address = items[5].getResponse();
  let phoneNumber = items[6].getResponse();
  let emailAddress = items[7].getResponse();
  let caseDetails = items[8].getResponse();

  // Find case number
  let caseNumber = '';
  const registrationLastRow = registrationSheet.getLastRow();
  const lastCaseNumber = parseInt(registrationSheet.getRange(registrationLastRow,1).getValue().slice(2));
  const lastCaseName = registrationSheet.getRange(registrationLastRow,3).getValue();
  if (lastCaseName.toLowerCase() === name.toLowerCase()) {
    caseNumber = 'RV' + (lastCaseNumber).toString().padStart(4, '0');
  } else {
    caseNumber = 'RV' + (lastCaseNumber + 1).toString().padStart(4, '0');
  }

  // Gendered responses
  let title = '';
  let sheHe = '';
  let herHim = '';
  let herHis = '';
  if (gender == 'Female') {
    title = 'Ms ';
    sheHe = 'she';
    herHim = 'her';
    herHis = 'her';
  } else if (gender == 'Male') {
    title = 'Mr ';
    sheHe = 'he';
    herHim = 'him';
    herHis = 'his';
  } else if (gender == 'Prefer not to say / Other') {
    title = '';
    sheHe = 'they';
    herHim = 'them';
    herHis = 'their';
  }

  // Get today's year and month
  const date = new Date();
  const month = Utilities.formatDate(date, 'GMT+8', 'MM');
  const year = Utilities.formatDate(date, 'GMT+8', 'yyyy');
  
  // Find and replace text in the letter body
  body.replaceText('{{Case_number}}', caseNumber);
  body.replaceText('{{Q}}', queueNumber);
  body.replaceText('{{Name}}', toTitleCase(name));
  body.replaceText('{{Name_Caps}}', name.toUpperCase());
  body.replaceText('{{NRIC}}', nric.toUpperCase());
  body.replaceText('{{Date_of_birth}}', dateOfBirth);  
  body.replaceText('{{Gender}}', gender);
  body.replaceText('{{Year}}', year);
  body.replaceText('{{Month}}', month);
  body.replaceText('{{Title}}', title);
  body.replaceText('{{she_he}}', sheHe);
  body.replaceText('{{her_him}}', herHim);
  body.replaceText('{{her_his}}', herHis);
  firstFooter.replaceText('{{Name}}', toTitleCase(name));
  firstFooter.replaceText('{{Address}}', nukeBlk(fixAddress(toTitleCase(address))));
  firstFooter.replaceText('{{Phone_number}}', phoneNumber);  
  firstFooter.replaceText('{{Email_address}}', emailAddress.toLowerCase());
  openDoc.saveAndClose(); // Save and close to flush updates and avoid weird errors

  // If case details provided input case details
  if (caseDetails != '') {
    openDoc2 = DocumentApp.openById(newTempFile.getId());
    body2 = openDoc2.getBody();
    body2.replaceText('{{Case_details}}', caseDetails);
    openDoc2.saveAndClose();
  }
  
  // Set the name
  const caseName = 'Agency' + year + month + caseNumber + '(subject)-' + toTitleCase(name);
  newTempFile.setName(caseName);

  // Move files to appropriate folder depending on whether case details are provided
  if (caseDetails != '') { // Move directly to 'Drafts' folder
    moveFiles(newTempFile.getId(), '1SB1Y_5P2Kc-oIPAvzeIs3aurqIpT4BzP');
    console.log('Created \'' + caseName + '\' in \'Drafts\'');
  } else { // Move to 'Registered' folder
    moveFiles(newTempFile.getId(), '1dsuxBMlKSjxJsAbrmMpzVKB-XhrOVIMA');
    console.log('Created \'' + caseName + '\' in \'Registered\'');
  }
}

function moveFiles(sourceFileId, targetFolderId) {
  let file = DriveApp.getFileById(sourceFileId);
  let folder = DriveApp.getFolderById(targetFolderId);
  file.moveTo(folder);
}

function toTitleCase(string, ignore=['a', 'an', 'and', 'at', 'but', 'by', 'for', 'in', 'nor', 'of', 'on', 'or', 'out', 'so', 'the', 'to', 'up', 'yet'], caps=['SKTC', 'S&CC', 'NRIC', 'HDB', 'BTO', 'SBF', 'MOP', 'CPF', 'MSF', 'SSO', 'FSC', 'ICA', 'PR', 'LTVP', 'STVP', 'EP', 'FDW', 'SPF', 'TP', 'LTA', 'PMD', 'TPE', 'KPE', 'SLE', 'CTE', 'LRT', 'MRT', 'MOH', 'SKH', 'ACE', 'MOM', 'WSG', 'TADM', 'TAFEP']) {
  ignore = new Set(ignore);
  caps = new Set(caps);
  return string.replace(/\w+/g, (word, i) => {
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

function fixAddress(str) { // Fixes alphanumeric block numbers like '182a Rivervale Crescent'
  return str.replace(
    /\d{1,4}[a-z]{1}\b/g,
    function(txt) {
      return txt.toUpperCase();
    }
  )
}

function nukeBlk(str) { // Removes the words 'Blk' or 'Block' in addresses because I hate it
  return str.replace(/\b[bB][lL]([oO][cC])?[kK]\s+\b/g, '');
}