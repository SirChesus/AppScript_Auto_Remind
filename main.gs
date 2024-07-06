// Get the active spreadsheet
const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

// Call the template from mail_template file

// Get the last row with data
const lastRow = sheet.getLastRow();
const range = sheet.getRange(2, 1, lastRow - 1, 5); // Adjusted to fetch 5 columns

// Function to get data from specific cells
function getInfo(row, col) {
  return sheet.getRange(row, col).getValue()
}

// Converts from milliseconds to days, then checks if (in days) <= limit
function isWithinRange(date, limit) {
  return (date - new Date()) / 8.64e+7 <= limit && date - new Date() > 0;
}

// Function which sends email when called
function sendMail(email, body) {
  var recEmail = String(email);
  console.log("HTML Body: ", String(body))

  if (!recEmail || recEmail == "") {
    console.error("Invalid email address or empty cell");
    return;
  }
  console.log("Check successful, email is:", recEmail);
  MailApp.sendEmail({
    to: recEmail,
    subject: "Peer Counseling Appointment Reminder",
    htmlBody: body
  });
}

// Main function indexes through the whole list, compares dates to decide if it should send an email
function main() {
  for (var i = 0; i < range.getNumRows(); i++) {
    let rowIndex = i + 2; // Skip header
    let apptDate = new Date(getInfo(rowIndex, 4)); 
    console.log("appt. date: ", apptDate)

    if (!isWithinRange(apptDate, 7) || apptDate == '') { // Assuming limit is 7 days
      continue;
    }
    const htmlBody = HtmlService.createTemplateFromFile('mail_template.html');
    htmlBody.name = getInfo(rowIndex,3); // Assuming column C is the name
    htmlBody.date = apptDate; // Assuming column D is the date
    htmlBody.period = getInfo(rowIndex, 5); // Assuming column E is the period

    const recipientEmail = getInfo(rowIndex, 2); // Assuming column B is the email
    sendMail(recipientEmail, htmlBody);
  }
}
