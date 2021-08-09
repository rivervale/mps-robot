function formResponse(e) {
  let respondentEmail = e.response.getRespondentEmail();
  let respondentName = e.response.getItemResponses()[0].getResponse();
  console.log('Respondent: ' + respondentName + '; Email: ' + respondentEmail);

  let mailSubject = 'Event registration confirmation: Sengkang Constituency Committee presents drowning prevention with Danny Ong from Project Silent';
  let mailSenderName = 'Sengkang Constituency Committee';
  let mailBody = `<p>Hello ${respondentName},</p>
  <p>Thank you for registering for our event, here are the event details. We look forward to seeing you soon!</p>
  <h3>Sengkang Constituency Committee presents: Drowning prevention with Danny Ong from Project Silent</h3>
  <p>
    <strong>Date:</strong> Thursday, 19 August 2021<br>
    <strong>Time:</strong> 8pm
  </p>
  <p>
    <strong>Zoom link:</strong> <a href='https://zoom.us/j/96194849373?pwd=R1JnSHl4V0o2UHZsbkQ5QmxsMWIxUT09'>https://zoom.us/j/96194849373?pwd=R1JnSHl4V0o2UHZsbkQ5QmxsMWIxUT09</a><br>
    <strong>Meeting ID:</strong> 961 9484 9373<br>
    <strong>Passcode:</strong> 456250
  </p>`;

  MailApp.sendEmail(respondentEmail, mailSubject, '', {
    name: mailSenderName,
    htmlBody: mailBody,
  });
}