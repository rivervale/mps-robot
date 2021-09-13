function checkConsultingFolder() {
  // This script is written in ES5 because of an error with the V8 engine when programatically creating time based triggers in Google Apps Script

  // Constants
  var folderIdConsulting =   '108hMH7ak3V2I8v5FkvMy2h7aB4QxShKb'; // Consulting folder
  var folderIdReadyToDraft = '1r-t8rdQv1SD2lmtHhHz0pFwAH2YKIk6R'; // Ready to draft folder
  var folderIdDrafts =       '1SB1Y_5P2Kc-oIPAvzeIs3aurqIpT4BzP'; // Drafts folder
  var folderIdArchive =      '1_JVQI0ZjCZ3MJwq1tsGrrTTFCeGr37PY'; // Case and document archive folder

  // Search for case sheet(s) with 'mpdone' in body but search can have a 5-10 min lag; the MP should input 'mpdone' when done with a case
  var diagnosedCases = DriveApp.getFolderById(folderIdConsulting).searchFiles("fullText contains 'mpdone'");
  // Alternatively the MP can shift the file into the 'Ready to Draft' folder which seems to work faster
  var diagnosedCasesAlt = DriveApp.getFolderById(folderIdReadyToDraft).getFilesByType('application/vnd.google-apps.document');

  // Create archive folder with today's date (if it does not already exist)
  var archiveFolderDate;
  if (diagnosedCases.hasNext() || diagnosedCasesAlt.hasNext()) {
    var today = Utilities.formatDate(new Date(), 'GMT+8', 'yyyyMMdd'); // Get today's date (yyyyMMdd)
    var archiveFolder = DriveApp.getFolderById(folderIdArchive);
    if (archiveFolder.getFoldersByName(today).hasNext()) {
      archiveFolderDate = archiveFolder.getFoldersByName(today).next(); // Get the folder with today's date if it exists
    } else {
      archiveFolderDate = archiveFolder.createFolder(today); // Otherwise create a folder with today's date
    }
  } else {
    // Exit the whole function if no completed cases found
    console.log('No completed cases found. Exiting');
    return;
  }

  while (diagnosedCases.hasNext()) {
    processCaseSheets(diagnosedCases, archiveFolderDate, folderIdDrafts);
  }

  while (diagnosedCasesAlt.hasNext()) {
    processCaseSheets(diagnosedCasesAlt, archiveFolderDate, folderIdDrafts);
  }
}

function processCaseSheets(fileIterator, archiveFolderDate, folderIdDrafts) {
  var caseSheet = fileIterator.next();
  var draftingTemplate = caseSheet.makeCopy(); // Create a copy of the case sheet to act as the drafting template

  var caseRef = caseSheet.getName(); // E.g. 'RV1000-202109-####: Tan Ah Seng'
  var caseRefTruncated = caseRef.slice(0, 13) + caseRef.slice(18); // E.g. 'RV1000-202109: Tan Ah Seng'

  moveFiles(caseSheet.getId(), archiveFolderDate.getId()); // Move original case sheet to the archive folder with today's date
  caseSheet.setName(caseRefTruncated); // Rename the case sheet to remove '-####' from the file name

  moveFiles(draftingTemplate.getId(), folderIdDrafts); // Move the drafting template to the 'Drafts' folder
  draftingTemplate.setName(caseRef); // Rename the drafting template to remove 'Copy of' from the file name

  console.log('Processed', caseRefTruncated);
}

// Run checkConsultingFolder() every 1 minutes
function periodicTrigger() {
  ScriptApp.newTrigger('checkConsultingFolder')
    .timeBased()
    .everyMinutes(1)
    .create();
}

function terminatePeriodicTrigger() {
  var triggers = getProjectTriggersByFunctionName('checkConsultingFolder');
  for (var i = 0; i < triggers.length; ++i) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

function getProjectTriggersByFunctionName(functionName) {
  return ScriptApp.getProjectTriggers().filter(function (trigger) {
    return trigger.getHandlerFunction() === functionName;
  });
}

function moveFiles(sourceFileId, targetFolderId) {
  var file = DriveApp.getFileById(sourceFileId);
  var folder = DriveApp.getFolderById(targetFolderId);
  file.moveTo(folder);
}

function weeklyTriggers() {
  // Run once to create the triggers
  // Start folder monitoring at 7.00pm on Mondays
  ScriptApp.newTrigger('periodicTrigger')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(19)
    .nearMinute(0)
    .inTimezone('Asia/Singapore')
    .create();

  // End folder monitoring at 11.45pm on Mondays
  ScriptApp.newTrigger('terminatePeriodicTrigger')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(23)
    .nearMinute(45)
    .inTimezone('Asia/Singapore')
    .create();
}