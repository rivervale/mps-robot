function checkConsultingFolder() {
  const folderIdConsulting =   '108hMH7ak3V2I8v5FkvMy2h7aB4QxShKb'; // Consulting folder
  const folderIdReadyToDraft = '1r-t8rdQv1SD2lmtHhHz0pFwAH2YKIk6R'; // Ready to draft folder
  const folderIdDrafts =       '1SB1Y_5P2Kc-oIPAvzeIs3aurqIpT4BzP'; // Drafts folder
  const folderIdArchive =      '1_JVQI0ZjCZ3MJwq1tsGrrTTFCeGr37PY'; // Case and document archive folder

  // Search for case sheet(s) with 'mpdone' in body but search can have a 5-10 min lag; the MP should input 'mpdone' when done with a case
  const diagnosedCases = DriveApp.getFolderById(folderIdConsulting).searchFiles("fullText contains 'mpdone'");
  // Alternatively the MP can shift the file into the 'Ready to Draft' folder which seems to work faster
  const diagnosedCasesAlt = DriveApp.getFolderById(folderIdReadyToDraft).getFilesByType('application/vnd.google-apps.document');

  // Create archive folder with today's date (if it does not already exist)
  if (diagnosedCases.hasNext() || diagnosedCasesAlt.hasNext()) {
    const today = Utilities.formatDate(new Date(), 'GMT+8', 'yyyyMMdd'); // Get today's date (yyyyMMdd)
    const archiveFolder = DriveApp.getFolderById(folderIdArchive);
    if (archiveFolder.getFoldersByName(today).hasNext()) {
      var archiveFolderDate = archiveFolder.getFoldersByName(today).next(); // Get the folder with today's date if it exists
    } else {
      var archiveFolderDate = archiveFolder.createFolder(today); // Otherwise create a folder with today's date
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
  let caseSheet = fileIterator.next();
  let draftingTemplate = caseSheet.makeCopy(); // Create a copy of the case sheet to act as the drafting template

  const caseRef = caseSheet.getName(); // E.g. 'RV1000-202109-####: Tan Ah Seng'
  const caseRefTruncated = caseRef.slice(0, 13) + caseRef.slice(18); // E.g. 'RV1000-202109: Tan Ah Seng'

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
  let triggers = getProjectTriggersByFunctionName('checkConsultingFolder');
  for (let i = 0; i < triggers.length; ++i) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

function getProjectTriggersByFunctionName(functionName) {
  return ScriptApp.getProjectTriggers().filter(
    (trigger) => trigger.getHandlerFunction() === functionName
  );
}

function moveFiles(sourceFileId, targetFolderId) {
  let file = DriveApp.getFileById(sourceFileId);
  let folder = DriveApp.getFolderById(targetFolderId);
  file.moveTo(folder);
}

function weeklyTriggers() {
  // Run once to create the triggers
  // Start folder monitoring at 6.30pm on Mondays
  ScriptApp.newTrigger('periodicTrigger')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(18)
    .nearMinute(30)
    .inTimezone('Asia/Singapore')
    .create();

  // End folder monitoring at 12.30am on Tuesdays
  ScriptApp.newTrigger('terminatePeriodicTrigger')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.TUESDAY)
    .atHour(0)
    .nearMinute(30)
    .inTimezone('Asia/Singapore')
    .create();
}
