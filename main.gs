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

// function which sends email when called
function sendMail(index, body){
  MailApp.sendEmail({
      
      to: String(getInfo('A',index)),
      subject: "Peer Counseling Appointment Reminder",
      htmlBody: body
      }); 

}

function main() {
  for(var i=0; i<range; i++){
    let newDate = new Date(Number(getInfo('D'+i)))
    let date = new Date()
    if(Number(newDate.getDate()-date.getDate()) == 2){
      htmlBody.name = getInfo('C'+ i)
      htmlBody.date = getInfo('D'+ i)
      htmlBody.period = getInfo('E'+ i) 
      var email_html = htmlBody.evaluate().getContent();
      sendMail(i, email_html)
    }
  }
}
