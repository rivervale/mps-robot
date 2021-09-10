function onFormSubmit(e) {
  // Inspired by: https://howtogapps.com/google-form-script-to-autofill-and-email-a-doc-template/

  // Folder and file IDs
  const folderIdRegistered =      '1dsuxBMlKSjxJsAbrmMpzVKB-XhrOVIMA'; // Registered folder
  const folderIdConsulting =      '108hMH7ak3V2I8v5FkvMy2h7aB4QxShKb'; // Consulting folder
  const folderIdDrafts =          '1SB1Y_5P2Kc-oIPAvzeIs3aurqIpT4BzP'; // Drafts folder
  const fileIdCaseSheetTemplate = '10taSkDqvqppcGIOPfvYWTpdBrROQmRsgkoA5eRDKwGc'; // Case sheet template
  const fileIdMpsData =           '1oUv4buU-IFAy9wqTDmdF_7eF40p8uTU_X8u16ujVKYU'; // Data spreadsheet

  // Case sheet template handling
  let templateDoc = DriveApp.getFileById(fileIdCaseSheetTemplate);
  let newTempFile = templateDoc.makeCopy(); //create a copy
  let openDoc = DocumentApp.openById(newTempFile.getId()); //open the new template document for editing
  let body = openDoc.getBody();
  let firstFooter = openDoc.getFooter().getParent().getChild(4);

  // Data spreadsheet handling
  let ss = SpreadsheetApp.openById(fileIdMpsData);
  const registrationSheet = ss.getSheetByName('Registration responses');
  
  // Get the responses triggered by onFormSubmit
  let items = e.response.getItemResponses();

  // Assign all form responses to variables
  let queueNumber = items[0].getResponse().padStart(2, '0'); //pads queue number to 2 digits
  let name = nukeAlias(items[1].getResponse());
  let nric = items[2].getResponse();
  let nricCensored = items[2].getResponse()[0] + '####' + items[2].getResponse().slice(5);
  let dateOfBirth = items[3].getResponse();
  let gender = items[4].getResponse();
  let address = items[5].getResponse();
  let phoneNumber = items[6].getResponse();
  let emailAddress = items[7].getResponse();
  let caseDetails = items[8].getResponse();

  // Find case numbers, including old cases
  Utilities.sleep(2000);
  let caseNumber = '';
  const nricRange = registrationSheet.getRange(2, 4, registrationSheet.getLastRow() - 1);
  const foundCases = nricRange.createTextFinder(nric).findAll();
  let caseNumbers = [];
  if (foundCases) {
    for (const foundCase of foundCases) {
      caseNumbers.push(registrationSheet.getRange(foundCase.getRow(), 1).getValue());
    }
  }
  caseNumber = caseNumbers.pop(); // Pop the last found casenumber from the array

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

  // Get dates
  const dateRaw = new Date();
  const dateYearMonth = Utilities.formatDate(dateRaw, 'GMT+8', 'yyyyMM');
  const dateFull = Utilities.formatDate(dateRaw, 'GMT+8', 'd MMMM yyyy');
  
  // Generate full case reference
  const caseRef = caseNumber + '-' + dateYearMonth + '-####';
  
  // Find and replace text in the letter body
  body.replaceText('{CaseReference}', caseRef);
  body.replaceText('{Date}', dateFull);
  body.replaceText('{CaseNumber}', caseNumber);
  body.replaceText('{Q}', queueNumber);
  body.replaceText('{Name}', toTitleCase(name).trim());
  body.replaceText('{NameCaps}', name.toUpperCase().trim());
  body.replaceText('{NRIC}', nricCensored.toUpperCase().trim());
  body.replaceText('{DOB}', dateOfBirth);
  body.replaceText('{PreviousCases}', (caseNumbers.length === 0 ? '' :' (prev. cases: ' + caseNumbers.toString() + ')'));
  body.replaceText('{Gender}', gender);
  body.replaceText('{Title}', title);
  body.replaceText('{SheHe}', sheHe);
  body.replaceText('{HerHim}', herHim);
  body.replaceText('{HerHis}', herHis);
  firstFooter.replaceText('{Name}', toTitleCase(name).trim());
  firstFooter.replaceText('{Address}', nukeBlk(fixAddress(toTitleCase(address))));
  firstFooter.replaceText('{PhoneNumber}', phoneNumber.trim());
  firstFooter.replaceText('{EmailAddress}', emailAddress.toLowerCase().trim());
  openDoc.saveAndClose(); // Save and close to flush updates and avoid weird errors

  // If case details provided input case details
  if (caseDetails != '') {
    let openDoc2 = DocumentApp.openById(newTempFile.getId());
    let body2 = openDoc2.getBody();
    body2.replaceText('{CaseDetails}', caseDetails);
    openDoc2.saveAndClose();
  }
  
  // Set the filename
  const caseName = caseRef + ': ' + toTitleCase(name);
  newTempFile.setName(caseName);

  // Move files to appropriate folder depending on whether case details are provided
  if (caseDetails != '') { // Move directly to 'Consulting' folder
    moveFiles(newTempFile.getId(), folderIdConsulting);
    console.log('Created \'' + caseName + '\' in \'Consulting\'');
  } else { // Move to 'Registered' folder
    moveFiles(newTempFile.getId(), folderIdRegistered);
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

function nukeAlias(str) { //Removes aliases from names as they are extraneous for our purposes
  return str.replace(/\b\s?((\((\s*\b\w+\b\s*)+\))|(@(\s*\b\w+\b\s*)+))/g, '');
}