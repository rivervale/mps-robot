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

function updateForm() {
  // https://howtogapps.com/create-an-issue-tracking-system-with-google-form-and-spreadsheet/

  // Open the MPS case writing form
  let form = FormApp.openById('1I4NSeB4iF74Hk9QYgAi8g7-BW2nspGcy65zyEJOKSSw');

  // Get the active spreadsheet
  let ss = SpreadsheetApp.getActive();

  // Get the case selection question
  let caseSelector = form.getItems()[0].asMultipleChoiceItem();

  // Extract all cases from this week and update the case selection question
  let weeksCasesSheet = ss.getSheetByName("Week's cases");
  let weeksCases = weeksCasesSheet.getRange(1, 1, weeksCasesSheet.getLastRow()).getValues();
  caseSelector.setChoiceValues(weeksCases);

  // Move template immediately if case details are included on registration
  Utilities.sleep(30000);
  let registrationSheet = ss.getSheetByName('Registration responses');
  let lastCaseDetails = registrationSheet.getRange(registrationSheet.getLastRow(),6).getValue();
  console.log('Last case details: ' + lastCaseDetails);
  if (lastCaseDetails != '') {
    let lastCaseName = toTitleCase(registrationSheet.getRange(registrationSheet.getLastRow(),3).getValue());
    console.log('Last case name: ' + lastCaseName);
    let matchingFiles = DriveApp.getFolderById('1dsuxBMlKSjxJsAbrmMpzVKB-XhrOVIMA').searchFiles("title contains '" + lastCaseName +"'");
    if (matchingFiles.hasNext()) {
      let date = new Date();
      let month = Utilities.formatDate(date, 'GMT+8', 'MM');
      let year = Utilities.formatDate(date, 'GMT+8', 'yyyy');
      let lastCaseNumber = registrationSheet.getRange(registrationSheet.getLastRow(),1).getValue();
      let workingDoc = matchingFiles.next();
      console.log('Found file: ' + workingDoc.getName());
      let openDoc = DocumentApp.openById(workingDoc.getId()); //open the doc for editing
      let body = openDoc.getBody();
      body.replaceText('{{Case_number}}', lastCaseNumber);
      workingDoc.setName('Agency' + year + month + lastCaseNumber + '(subject)-' + lastCaseName);
      moveFiles(workingDoc.getId(), '1SB1Y_5P2Kc-oIPAvzeIs3aurqIpT4BzP');
    } else {
      console.log('Failed to find file');
    }
  }
}