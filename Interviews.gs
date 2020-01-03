var EMAIL_DRAFTED = "EMAIL DRAFTED";
var EMAIL_SENT = "EMAIL SENT";

function interviewEmails() {
  var sheet = SpreadsheetApp.getActiveSheet(); // Use data from the active sheet
  var startRow = 2;                            // First row of data to process
  var numRows = sheet.getLastRow() - 1;        // Number of rows to process
  var lastColumn = sheet.getLastColumn();      // Last column
  var dataRange = sheet.getRange(startRow, 1, numRows, lastColumn) // Fetch the data range of the active sheet
  var data = dataRange.getValues();            // Fetch values for each row in the range
  
  // Work through each row in the spreadsheet
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];  
    // Assign each row a variable
    var clientName = row[1];                // Col A: Client name
    var clientEmail = row[0];               // Col B: Client email
    var date = row[3];                      
    var time = Utilities.formatDate(row[4], SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "h:mm a");
    var room = row[5];                   
    var emailStatus = row[lastColumn - 1];  
    
    // Prevent from drafing duplicates and from drafting emails without a recipient
    if (emailStatus !== EMAIL_DRAFTED && clientEmail) {  
    
      // Build the email message
      var htmlOutput = HtmlService.createHtmlOutputFromFile('Email_Interview')
      var message = htmlOutput.getContent()
      message = message.replace("%name", clientName);
      message = message.replace("%date", date);
      message = message.replace("%time", time);
      message = message.replace("%room", room);
          
      
      // Create the email draft
      GmailApp.sendEmail(
        clientEmail,            // Recipient
        '[TEST] VSA Intern Interview',  // Subject
        '',                     // Body (plain text)
        {
        htmlBody: message    // Options: Body (HTML)
        }
      );
      
      sheet.getRange(startRow + i, lastColumn).setValue(EMAIL_DRAFTED); // Update the last column with "EMAIL_DRAFTED"
      SpreadsheetApp.flush(); // Make sure the last cell is updated right away
    }
  }
}
