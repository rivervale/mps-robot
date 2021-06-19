function updateForm() {
  // Inspired by: https://howtogapps.com/create-an-issue-tracking-system-with-google-form-and-spreadsheet/

  // Open the MPS case writing form
  let form = FormApp.openById('1I4NSeB4iF74Hk9QYgAi8g7-BW2nspGcy65zyEJOKSSw');

  // Get the active spreadsheet
  let ss = SpreadsheetApp.getActive();

  // Get all relevant sheets
  let selfRegisteredSheet = ss.getSheetByName('Self-registered cases');
  let registrationSheet = ss.getSheetByName('Registration responses');
  let caseWriterSheet = ss.getSheetByName('Case writing responses');
  let weeksCasesSheet = ss.getSheetByName('Week\'s cases');

  // Get the case selection question in the case writing form
  let caseSelector = form.getItems()[0].asMultipleChoiceItem();

  // Extract all cases from this week and update the case selection question in the case writing form
  let weeksCases = weeksCasesSheet.getRange(1, 1, weeksCasesSheet.getLastRow()).getValues();
  caseSelector.setChoiceValues(weeksCases);

  // Resize last two rows of spreadsheets to 53 pixels high for neatness
  selfRegisteredSheet.setRowHeightsForced(selfRegisteredSheet.getLastRow() - 1, 2, 53);
  registrationSheet.setRowHeightsForced(registrationSheet.getLastRow() - 1, 2, 53);
  caseWriterSheet.setRowHeightsForced(caseWriterSheet.getLastRow(), 1, 53);
}

function moveFiles(sourceFileId, targetFolderId) {
  let file = DriveApp.getFileById(sourceFileId);
  let folder = DriveApp.getFolderById(targetFolderId);
  file.moveTo(folder);
}

function toTitleCase(string, ignore=['a', 'an', 'and', 'at', 'but', 'by', 'for', 'in', 'nor', 'of', 'on', 'or', 'out', 'so', 'the', 'to', 'up', 'yet']) {
  ignore = new Set(ignore);
  return string.replace(/\w+/g, (word, i) => {
    word = word.toLowerCase();
    if (i && ignore.has(word)) {
      return word;
    }
    return word[0].toUpperCase() + word.slice(1);
  });
};