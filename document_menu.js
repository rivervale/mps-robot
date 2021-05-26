let DOCUMENT = DocumentApp.getActiveDocument().getId();

function onOpen() {
  let ui = DocumentApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('MPS')
      .addItem('Move to folder for SA\'s review', 'move1')
      .addItem('Move to folder for MP\'s review', 'move2')
      .addSeparator()
      .addItem('Mark as \'Ready to email\'', 'move3')
      .addItem('Mark as \'Ready to post\'', 'move4')
      .addItem('Mark as \'Sent\'', 'move5')
/*    .addSeparator()
      .addSubMenu(ui.createMenu('Move')
          .addItem('Move to SA review', 'move1')
          .addItem('Move to MP review', 'move2'))*/
      .addToUi();
}

function moveFiles(sourceFileId, targetFolderId) {
  let file = DriveApp.getFileById(sourceFileId);
  let folder = DriveApp.getFolderById(targetFolderId);
  file.moveTo(folder);
}

function move1() { // Move to SA review folder
  moveFiles(DOCUMENT, '1qYWTyZ4MqivXVVQEJRXwpGvcBKWz1B4b');
}

function move2() { // Move to MP review folder
  moveFiles(DOCUMENT, '1I3cBU00RCeLow-KyamlStso2-zGV3VjP');
}

function move3() { // Move to 'Ready to email' folder
  moveFiles(DOCUMENT, '1oV0J-u7AWjwxnByegeNf3JJpOBoBs7_T');
}

function move4() { // Move to 'Ready to post' folder
  moveFiles(DOCUMENT, '16ZAXAFxaMS30VOoZSGqZHheL_MpeWoq_');
}

function move5() { // Move to 'Sent' folder
  moveFiles(DOCUMENT, '1EFxENHZJSFoLdBlg-j-Zyu57JBpoA-k9');
}