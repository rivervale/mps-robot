function onFormSubmit(e) {
  // Get the responses triggered by onFormSubmit
  let items = e.response.getItemResponses();

  // Assign all form responses to variables
  let name = toTitleCase(items[0].getResponse().slice(32));
  let caseDetails = items[1].getResponse();

  // Document handling
  let files = DriveApp.getFolderById('1dsuxBMlKSjxJsAbrmMpzVKB-XhrOVIMA').searchFiles('title contains "' + name.replace(/'/g, '\'') +'"'); // .replace function helps to escape names with single quotes
  let workingDoc = files.next();
  let openDoc = DocumentApp.openById(workingDoc.getId()); //open the doc for editing
  let body = openDoc.getBody();
  
  // Insert case details into existing letter body
  body.replaceText('{{Case_details}}', caseDetails);
  
  // Save and close the open document and move it to 'Drafts' folder
  openDoc.saveAndClose();
  moveFiles(workingDoc.getId(), '1SB1Y_5P2Kc-oIPAvzeIs3aurqIpT4BzP');
  console.log('Updated \'' + workingDoc.getName() + '\' with case details');
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