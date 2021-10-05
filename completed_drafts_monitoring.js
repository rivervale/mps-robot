function completedDraftsMonitoring() {
  // Constants
  const folderIdDrafts = '1SB1Y_5P2Kc-oIPAvzeIs3aurqIpT4BzP'; // Drafts folder
  const folderIdSaReview = '1qYWTyZ4MqivXVVQEJRXwpGvcBKWz1B4b'; // SA Review folder
  const folderIdExclusions = ['1HAMYLzj9tP4zK0IqCLppPWREBkwOn2Jt']; // Array of folder IDs to exclude from monitoring
  const emailAddressSA = 'andrelowwy@gmail.com';
  const saFolderUrl = `https://drive.google.com/drive/u/0/folders/${folderIdSaReview}`;

  let counter = 0;

  // Rolling log to be emailed at the end after execution
  let mailLog = `<h1>MPS Robot: Draft(s) completed</h1>
  <p>The following case(s) are written and ready to review:</p>
  <table border='1' style='border-collapse: collapse;'>
    <tr>
      <th scope="col">Case</th>
      <th scope="col">Resident</th>
      <th scope="col">Agency</th>
      <th scope="col">Drafted by</th>
      <th scope="col">Link</th>
    </tr>
  `;

  // Find all child folders in the Drafts folder
  let draftsFolderChildren =
    DriveApp.getFolderById(folderIdDrafts).getFolders();

  while (draftsFolderChildren.hasNext()) {
    let currentFolder = draftsFolderChildren.next();

    // Check if the current folder matches exclusions
    if (folderIdExclusions.includes(currentFolder.getId())) {
      continue; // Skips current loop if folder matches exclusions
    }

    // Find all Google Docs files in current folder and process them
    let currentFiles = currentFolder.getFilesByType(
      'application/vnd.google-apps.document'
    );
    while (currentFiles.hasNext()) {
      let processOutput = processFiles(
        currentFiles,
        currentFolder,
        folderIdSaReview
      );

      if (processOutput) {
        counter += 1;
        // Update mailLog
        mailLog += `<tr>
          <td>${processOutput[0]}</td>
          <td>${processOutput[1]}</td>
          <td>${processOutput[2]}</td>
          <td>${processOutput[3]}</td>
          <td><a href='${processOutput[4]}'>Open</a></td>
        </tr>
        `;
      }
    }
  }

  // Send a log if any cases were processed
  if (counter > 0) {
    // Close off the mailLog HTML
    mailLog += `</table>
      <p>
        <a href='${saFolderUrl}'
          style='
            background-color: #007FFF;
            border: 1px solid #007FFF;
            border-radius: 5px;
            color: white;
            padding: 10px 20px;
            text-align: center;
            text-decoration: none;
            display: inline-block;
            font-size: 16px;'>
          Review folder
        </a>
      </p>
      `;

    // Send a log to the SA
    MailApp.sendEmail({
      name: 'MPS Robot',
      subject: 'Draft(s) completed',
      to: emailAddressSA,
      htmlBody: mailLog,
    });
  }
}

function processFiles(fileIterator, currentFolder, saReviewFolder) {
  // Reusable function to process files to see if they are completed and can be moved to 'SA Review' folder
  const inProgressKeyphrase = 'MP’s actions'; // Keyphrase to detect if work still in progress

  let currentFile = fileIterator.next();
  let fileName = currentFile.getName();
  let caseNumber = currentFile.getName().slice(0, 6);
  let resident = currentFile.getName().split(': ')[1];
  let agency = currentFile.getName().slice(14).split(':')[0];
  let drafter = currentFolder.getName();
  let fileUrl = currentFile.getUrl();

  // Check if case is in progress
  if (
    DocumentApp.openById(currentFile.getId())
      .getBody()
      .findText(inProgressKeyphrase)
  ) {
    return false; // If so, the function can end here; no need to move to 'SA Review' folder
  }

  // If the case is completed, it can be moved to the 'SA Review' folder
  moveFiles(currentFile.getId(), saReviewFolder);
  console.log(drafter, 'completed', fileName, '(moved for review)');
  return [caseNumber, resident, agency, drafter, fileUrl];
}

function moveFiles(sourceFileId, targetFolderId) {
  // Reusable function to move files
  let file = DriveApp.getFileById(sourceFileId);
  let folder = DriveApp.getFolderById(targetFolderId);
  file.moveTo(folder);
}