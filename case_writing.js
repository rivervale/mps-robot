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

function onFormSubmit(e) {
  //https://howtogapps.com/google-form-script-to-autofill-and-email-a-doc-template/
  
  //Get today's year and month
  let date = new Date();
  let month = Utilities.formatDate(date, 'GMT+8', 'MM');
  let year = Utilities.formatDate(date, 'GMT+8', 'yyyy');

  //Get the responses triggered by On Form Submit
  let items = e.response.getItemResponses();

  //assign form responses to variables
  //items[0].getResponse() is the first response in the Form
  let caseNumber = items[0].getResponse().slice(0,6);
  let name = toTitleCase(items[0].getResponse().slice(32));
  let caseDetails = items[1].getResponse();

  //Document handling
  let files = DriveApp.getFolderById('1dsuxBMlKSjxJsAbrmMpzVKB-XhrOVIMA').searchFiles("title contains '" + name +"'");
  let workingDoc = files.next();
  let openDoc = DocumentApp.openById(workingDoc.getId()); //open the doc for editing
  let body = openDoc.getBody();
  
  //find the and replace text in the template
  body.replaceText('{{Case_number}}', caseNumber);
  body.replaceText('{{Case_details}}', caseDetails);
  
  //Save and Close the open document and set the name
  openDoc.saveAndClose();
  workingDoc.setName('Agency' + year + month + caseNumber + '(subject)-' + toTitleCase(name));
  moveFiles(workingDoc.getId(), '1SB1Y_5P2Kc-oIPAvzeIs3aurqIpT4BzP');
}