function onFormSubmit(e) {
  // Folder IDs
  const folderIdRegistered = '1dsuxBMlKSjxJsAbrmMpzVKB-XhrOVIMA'; // Registered folder
  const folderIdConsulting = '108hMH7ak3V2I8v5FkvMy2h7aB4QxShKb'; // Consulting folder
  const folderIdDrafts =     '1SB1Y_5P2Kc-oIPAvzeIs3aurqIpT4BzP'; // Drafts folder

  // Get the responses triggered by onFormSubmit
  let items = e.response.getItemResponses();

  // Assign all form responses to variables
  const caseNumber = items[0].getResponse().slice(0,6); // Extracts the case number (first 6 characters) from a string with format resembling 'RV1000; Q: 01; ID: —123A; Name: Tan'
  const caseDetails = items[1].getResponse();
  const caseWriter = items[2].getResponse();

  // Searching for the case sheet
  const files = DriveApp.getFolderById(folderIdRegistered).searchFiles("title contains '" + caseNumber +"'"); // Search the 'Registered' folder for the case sheet by case number (e.g. 'RV1000')
  console.log('Searching for case:', caseNumber);
  try {
    var workingDoc = files.next(); // Select the first matching case sheet
  } catch (error) { // Error handling if no matching case sheet found
    console.log('Case not found:', caseNumber);
    console.log(error);
    return; // End function
  }

  // Document handling
  let openDoc = DocumentApp.openById(workingDoc.getId()); // Open the case sheet for editing
  let body = openDoc.getBody();
  
  // Insert case details into existing letter body
  body.replaceText('{CaseWriter}', ' (Written by: ' + caseWriter + ')');
  body.replaceText('{CaseDetails}', caseDetails);
  
  // Save and close the open document and move it to 'Consulting' folder
  openDoc.saveAndClose();
  moveFiles(workingDoc.getId(), folderIdConsulting);
  console.log('Updated with case details:', workingDoc.getName());
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