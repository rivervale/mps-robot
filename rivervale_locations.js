function updateForm() {
  const houseVisitForm = '1RQck_N5atbVxsPz-YVOyCCaCERkraEsh5PMbdmka1_0'; // ID for house visit form

  let form = FormApp.openById(houseVisitForm); // Open the house visit form

  let spreadsheet = SpreadsheetApp.getActive(); // Get the active spreadsheet
  let rivervaleLocationsSheet = spreadsheet.getSheetByName('Rivervale locations')

  let locationSelector = form.getItems()[0].asListItem();

  let rivervaleLocations = rivervaleLocationsSheet.getRange(1,1,rivervaleLocationsSheet.getLastRow()).getValues();

  locationSelector.setChoiceValues(rivervaleLocations);
}