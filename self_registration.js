function emailOnSubmit(e) {
  // Get the responses triggered by onFormSubmit
  let items = e.response.getItemResponses();

  // Assign all form responses to variables
  let name = toTitleCase(items[0].getResponse().trim());
  let nric = items[1].getResponse();
  let nricCensored =
    items[1].getResponse()[0] + '####' + items[1].getResponse().slice(5);
  let address = fixAddress(toTitleCase(items[2].getResponse().trim()));
  let gender = items[3].getResponse();
  let phoneNumber = items[4].getResponse();
  let emailAddress = items[5].getResponse().toLowerCase();
  let repeatCustomer = items[6].getResponse().slice(0, 3);
  let caseDetails = items[7].getResponse().trim();
  let supportingDocs = items[8].getResponse().slice(0, 3);

  // Link to spreadsheet with case responses
  const caseResponsesUrl =
    'https://docs.google.com/spreadsheets/d/1oUv4buU-IFAy9wqTDmdF_7eF40p8uTU_X8u16ujVKYU/edit#gid=1588694449';

  // Generate case acceptance link for use later
  const acceptCaseUrl = `https://docs.google.com/forms/d/e/1FAIpQLSfF6b96fzmTvVrSEcR_iDnp-eYhcTBZYdwSYxv-FtldchdyMQ/viewform?usp=pp_url&entry.259633438=${name.replace(
    / /g,
    '+'
  )}&entry.496077513=${nric}&entry.1209782293=${gender.replace(
    / /g,
    '+'
  )}&entry.529557127=${fixedEncodeURIComponent(address).replace(
    /%20/g,
    '+'
  )}&entry.1922153968=${phoneNumber}&entry.1807123632=${emailAddress}&entry.501929772=${fixedEncodeURIComponent(
    caseDetails
  ).replace(/%20/g, '+')}`;

  // Email a summary when someone fills in the form
  let mailRecipients = 'khengwee.chua@wp.sg';
  let mailRecipientsBcc = 'andrelowwy@gmail.com, rivervale.wp@gmail.com';
  let mailSubject = 'New MPS customer: ' + name;
  let mailSenderName = 'MPS Robot';
  let mailSenderReplyTo = 'khengwee.chua@wp.sg';
  let mailBody = `
    <h1>MPS case submitted</h1>
    <p>${name} has submitted a case for consideration.</p>
    <h2>Personal details</h2>
    <table style='border: none;'>
      <tr>
        <td><strong>Name:</strong></td>
        <td>${name}</td>
      </tr>
      <tr>
        <td><strong>NRIC:</strong></td>
        <td>${nricCensored}</td>
      </tr>
      <tr>
        <td><strong>Gender:</strong></td>
        <td>${gender}</td>
      </tr>
      <tr>
        <td><strong>Address:</strong></td>
        <td style="white-space: pre-line">${address}</td>
      </tr>
      <tr>
        <td><strong>Phone:</strong></td>
        <td>${phoneNumber}</td>
      </tr>
      <tr>
        <td><strong>Email:</strong></td>
        <td>${emailAddress}</td>
      </tr>
    </table style='border: none;'>
    <h2>Case details</h2>
    <table>
      <tr>
        <td><strong>Details:</strong></td>
        <td style="white-space: pre-line">${caseDetails}</td>
      </tr>
      <tr>
        <td><strong>Repeat case?:</strong></td>
        <td>${repeatCustomer}</td>
      </tr>
      <tr>
        <td><strong>Any docs?:</strong></td>
        <td>${supportingDocs}</td>
      </tr>
    </table>
    <p>
      <a href='${caseResponsesUrl}'
        style='
          background-color: white;
          border: 1px solid #007FFF;
          border-radius: 5px;
          color: #007FFF;
          padding: 10px 20px;
          text-align: center;
          text-decoration: none;
          display: inline-block;
          font-size: 16px;'>
        See responses
      </a>
      <a href='${acceptCaseUrl}'
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
        Accept case
      </a>
    </p>
  `;

  MailApp.sendEmail(mailRecipients, mailSubject, '', {
    name: mailSenderName,
    bcc: mailRecipientsBcc,
    htmlBody: mailBody,
    replyTo: mailSenderReplyTo,
  });

  console.log('New MPS customer: ' + name);
}

function toTitleCase(str) {
  return str.replace(/\w\S*/g, function (txt) {
    return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
  });
}

function fixAddress(str) {
  // Fixes block numbers like '182a Rivervale Crescent'
  return str.replace(/\d{1,4}[a-z]{1}\b/g, function (txt) {
    return txt.toUpperCase();
  });
}

function fixedEncodeURIComponent(str) {
  return encodeURIComponent(str).replace(/[!'()*]/g, function (c) {
    return '%' + c.charCodeAt(0).toString(16);
  });
}
