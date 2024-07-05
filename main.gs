// Get the active spreadsheet
const sheet = SpreadsheetApp.getActiveSheet();

// Call the template from mail_template file
const htmlBody = HtmlService.createTemplateFromFile('mail_template')

// Get the last row with data EDIT
const lastRow = sheet.getLastRow();
const range = sheet.getRange(2, 1, lastRow - 1, 5); // Adjusted to fetch 5 columns

// Function to get data from specific cells
function getInfo(col, index) {
  return sheet.getRange(col + index).getValue();
}

// converts from milliseconds to days, then checks if (in days) <= limit
function isWithinRange(date = new Date, limit){
  return (date - new Date())/8.64e+7 <= limit && date - new Date() > 0
}

// Function which sends email when called
function sendMail(email, body) {
  var recipientEmail = String(email)

  if (!recipientEmail || recipientEmail === "") {
        console.error("Invalid email address or empty cell at row " + rowIndex);
        return;
      } 
  MailApp.sendEmail({
    // Error in the recipientEmail, body is not being sent properly aswell
    to: 'mayankatte4707@student.lvusd.org',
    subject: "Peer Counseling Appointment Reminder",
    htmlBody: body
  });
}

// Main function indexes through whole list, compares dates to decide if it should send an email
function main() {
  for (var i = 0; i < range.getNumRows(); i++) {
    let rowIndex = i + 2; // Skip headers

    let apptDate = new Date(getInfo("A", rowIndex)); // Assuming 'D' is column 4 (index 3)

    // Compare dates considering year, month, and day
    if (isWithinRange(getInfo(apptDate, 2))) {

      htmlBody.name = getInfo("C", rowIndex); 
      htmlBody.date = getInfo("D", rowIndex); 
      htmlBody.period = getInfo("E", rowIndex); 

      var recipientEmail = getInfo('B', rowIndex); 
      var email_html = htmlBody.evaluate().getContent();
      sendMail(recipientEmail, email_html);
       
      }
    }
  }
