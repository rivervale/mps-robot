function checkConsultingFolder() {
  // This script is written in ES5 because of an error with the V8 engine when programatically creating time based triggers in Google Apps Script

  // Constants
  var folderIdConsulting =   '108hMH7ak3V2I8v5FkvMy2h7aB4QxShKb'; // Consulting folder
  var folderIdReadyToDraft = '1r-t8rdQv1SD2lmtHhHz0pFwAH2YKIk6R'; // Ready to draft folder
  var folderIdDrafts =       '1SB1Y_5P2Kc-oIPAvzeIs3aurqIpT4BzP'; // Drafts folder
  var folderIdArchive =      '1_JVQI0ZjCZ3MJwq1tsGrrTTFCeGr37PY'; // Case and document archive folder

  // Scan the 'Ready to Draft' sub-folder for files
  var diagnosedCases = DriveApp.getFolderById(folderIdReadyToDraft).getFilesByType('application/vnd.google-apps.document');
  var supportingDocsPdf = DriveApp.getFolderById(folderIdReadyToDraft).getFilesByType('application/pdf');
  var supportingDocsJpeg = DriveApp.getFolderById(folderIdReadyToDraft).getFilesByType('image/jpeg');
  var supportingDocsPng = DriveApp.getFolderById(folderIdReadyToDraft).getFilesByType('image/png');

  // Create archive folder with today's date (if it does not already exist)
  var archiveFolderDate;
  if (diagnosedCases.hasNext()) {
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

  // While each file iterator method has detected files, process the case sheets
  while (diagnosedCases.hasNext()) {
    processCaseSheets(diagnosedCases, archiveFolderDate, folderIdDrafts);
  }

  // While supporting docs remaing, process supporting docs
  while (supportingDocsPdf.hasNext()) {
    processSupportingDocs(supportingDocsPdf, archiveFolderDate, folderIdDrafts);
  }
  while (supportingDocsJpeg.hasNext()) {
    processSupportingDocs(supportingDocsJpeg, archiveFolderDate, folderIdDrafts);
  }
  while (supportingDocsPng.hasNext()) {
    processSupportingDocs(supportingDocsPng, archiveFolderDate, folderIdDrafts);
  }
}

function processCaseSheets(fileIterator, archiveFolder, draftsFolder) {
  // Reusable function to move case sheets to archive folder and create a drafting template in the drafts folder

  // Constants
  var nfaKeyword = 'NFA'; // Keyword to mark case as NFA (i.e. no further action required)

  var caseSheet = fileIterator.next();

  var caseRef = caseSheet.getName(); // E.g. 'RV1000-202109-####: Tan Ah Seng'
  var caseRefTruncated = caseRef.slice(0, 13) + caseRef.slice(18); // E.g. 'RV1000-202109: Tan Ah Seng'

  moveFiles(caseSheet.getId(), archiveFolder.getId()); // Move original case sheet to the archive folder with today's date
  caseSheet.setName(caseRefTruncated); // Rename the case sheet to remove '-####' from the file name

  if (DocumentApp.openById(caseSheet.getId()).getBody().findText(nfaKeyword)) { // Check if case is of the 'No Further Action' variety
    console.log ('Processed', caseRefTruncated, '(no further action)')
    return; // If so, the function can end here; no need to create a drafting template
  }

  var draftingTemplate = caseSheet.makeCopy(); // Create a copy of the case sheet to act as the drafting template
  moveFiles(draftingTemplate.getId(), draftsFolder); // Move the drafting template to the 'Drafts' folder
  draftingTemplate.setName(caseRef); // Rename the drafting template to remove 'Copy of' from the file name

  console.log('Processed', caseRefTruncated);
}

function processSupportingDocs(fileIterator, archiveFolder, draftsFolder) {
  // Reusable function to move supporting docs to archive folder and create a copy in the drafts folder

  var supportingDoc = fileIterator.next();

  var supportingDocRef = supportingDoc.getName().slice(0, 6) + ' docs';

  moveFiles(supportingDoc.getId(), archiveFolder.getId()); // Move original supporting doc to the archive folder with today's date
  var supportingDocCopy = supportingDoc.makeCopy(); // Create a copy of the supporting doc
  moveFiles(supportingDocCopy.getId(), draftsFolder); // Move the copied supporting doc to the 'Drafts' folder
  supportingDocCopy.setName(supportingDocRef); // Rename the copied supporting doc to remove 'Copy of' from the file name

  console.log('Processed supporting doc', supportingDocRef);
}

function periodicTrigger() {
  // Run checkConsultingFolder() every 1 minutes
  ScriptApp.newTrigger('checkConsultingFolder')
    .timeBased()
    .everyMinutes(1)
    .create();
}

function terminatePeriodicTrigger() {
  // Terminate any periodic triggers running checkConsultingFolder()
  var triggers = getProjectTriggersByFunctionName('checkConsultingFolder');
  for (var i = 0; i < triggers.length; ++i) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

function getProjectTriggersByFunctionName(functionName) {
  // Resuable function that returns triggers given a function that it triggers
  return ScriptApp.getProjectTriggers().filter(function (trigger) {
    return trigger.getHandlerFunction() === functionName;
  });
}

function moveFiles(sourceFileId, targetFolderId) {
  // Reusable function to move files
  var file = DriveApp.getFileById(sourceFileId);
  var folder = DriveApp.getFolderById(targetFolderId);
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

  // End folder monitoring at 11.30pm on Mondays
  ScriptApp.newTrigger('terminatePeriodicTrigger')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(23)
    .nearMinute(30)
    .inTimezone('Asia/Singapore')
    .create();
}