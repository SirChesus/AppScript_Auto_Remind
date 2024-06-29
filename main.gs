// get the active spreadsheet
const sheet = SpreadsheetApp.getActiveSheet();

// call the template from mail_template file
const htmlBody = HtmlService.createTemplateFromFile('mail_template');

// rows to get length later
const lastRow = sheet.getLastRow();
const range = sheet.getRange(2, 1, lastRow - 1, 3); // Row index set to 2 to skip header row

// gets whatever is in 'col' column at 'index' index
function getInfo(col, index){
  return sheet.getRange(col+index).getValue()
}

function formatTime(index){
  //code
}

// takes the time from now to future date and converts into days, then checks if < limit and > 0 (idk if needed)
function checkWithinRange(date = new Date, limit){
  return (date - new Date())/8.64e+7 <= limit && date - new Date() > 0
}

// no clue why this is needed, just keep it
var email_html = htmlBody.evaluate().getContent();

// function which sends email when called
function sendMail(index){
  MailApp.sendEmail({
      
      to: String(getInfo('A',index)),
      subject: "Peer Counseling Appointment Reminder",
      htmlBody: email_html
      }); 

}

function main() {
  for(var i=0; i<range; i++){
    let newDate = new Date(Number(getInfo("D"+i)))
    let date = new Date()
    if(Number(newDate.getDate()-date.getDate()) == 2){
      sendMail(i)
    }
  }
}
