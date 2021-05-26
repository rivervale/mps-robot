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

  //Template handling
  let templateDoc = DriveApp.getFileById('10taSkDqvqppcGIOPfvYWTpdBrROQmRsgkoA5eRDKwGc');
  let newTempFile = templateDoc.makeCopy(); //create a copy
  let openDoc = DocumentApp.openById(newTempFile.getId()); //open the new template document for editing
  let body = openDoc.getBody();
  let firstFooter = openDoc.getFooter().getParent().getChild(4);
  
  //Get the responses triggered by On Form Submit
  let items = e.response.getItemResponses();

  //assign all form responses to variables
  //items[0].getResponse() is the first response in the Form
  let queueNumber = items[0].getResponse().padStart(2, '0'); //pads queue number to 2 digits
  let name = items[1].getResponse();
  let nric = items[2].getResponse()[0] + '####' + items[2].getResponse().slice(5);
  let dateOfBirth = items[3].getResponse();
  let gender = items[4].getResponse();
  let address = items[5].getResponse();
  let phoneNumber = items[6].getResponse();
  let emailAddress = items[7].getResponse();
  let caseDetails = items[8].getResponse();

  //gendered responses
  let title = '';
  let sheHe = '';
  let herHim = '';
  let herHis = '';
  if (gender == 'Female') {
    title = 'Ms ';
    sheHe = 'she';
    herHim = 'her';
    herHis = 'her';
  } else if (gender == 'Male') {
    title = 'Mr ';
    sheHe = 'he';
    herHim = 'him';
    herHis = 'his';
  } else if (gender == 'Prefer not to say / Other') {
    title = '';
    sheHe = 'they';
    herHim = 'them';
    herHis = 'their';
  }
  
  //find and replace text in the letter body
  body.replaceText('{{Q}}', queueNumber);
  body.replaceText('{{Name}}', toTitleCase(name));
  body.replaceText('{{Name_Caps}}', name.toUpperCase());
  body.replaceText('{{NRIC}}', nric.toUpperCase());
  body.replaceText('{{Date_of_birth}}', dateOfBirth);  
  body.replaceText('{{Gender}}', gender);
  body.replaceText('{{Year}}', year);
  body.replaceText('{{Month}}', month);
  body.replaceText('{{Title}}', title);
  body.replaceText('{{she_he}}', sheHe);
  body.replaceText('{{her_him}}', herHim);
  body.replaceText('{{her_his}}', herHis);
  if (caseDetails != "") {
    body.replaceText('{{Case_details}}', caseDetails);
  }
  firstFooter.replaceText('{{Name}}', toTitleCase(name));
  firstFooter.replaceText('{{Address}}', toTitleCase(address));
  firstFooter.replaceText('{{Phone_number}}', phoneNumber);  
  firstFooter.replaceText('{{Email_address}}', emailAddress.toLowerCase());
  
  //Save and Close the open document and set the name
  openDoc.saveAndClose();
  newTempFile.setName('Agency' + year + month + 'RV####(subject)-' + toTitleCase(name));
  moveFiles(newTempFile.getId(), '1dsuxBMlKSjxJsAbrmMpzVKB-XhrOVIMA');
}