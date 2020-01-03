var EMAIL_DRAFTED = "EMAIL DRAFTED";

function draftMyEmails() {
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
    var apples = row[3];                       // Col C: Vegetable name
    var bananas = row[4];                   // Col D: Vegetable description
    var cherries = row[5];  // Col E: Email Status
    var emailStatus = row[lastColumn - 1];  // Col E: Email Status
    
    // Prevent from drafing duplicates and from drafting emails without a recipient
    if (emailStatus !== EMAIL_DRAFTED && clientEmail) {  
    
      // Build the email message
      var emailBody =  '<p>Hi ' + clientName + ',<p>';
          emailBody += '<p>We are pleased to match you with your vegetable: <strong>' + row[1] + '</strong><p>';
          emailBody += '<h2>About ' + clientEmail + '</h2>';
          emailBody += '<p>You ordered: \n </p>';
          emailBody += '<p>x' + apples + ' Apples</p>';
          emailBody += '<p>x' + bananas + ' Bananas</p>';
          emailBody += '<p>x' + cherries + ' Cherries</p>';
          emailBody += '<p>' + clientName + ', we hope that you and ' + cherries + ' have a wonderful relationship.<p>';
          
      
      // Create the email draft
      GmailApp.createDraft(
        clientEmail,            // Recipient
        'Fruit Confirmation',  // Subject
        '',                     // Body (plain text)
        {
        htmlBody: emailBody    // Options: Body (HTML)
        }
      );
      
      sheet.getRange(startRow + i, lastColumn).setValue(EMAIL_DRAFTED); // Update the last column with "EMAIL_DRAFTED"
      SpreadsheetApp.flush(); // Make sure the last cell is updated right away
    }
  }
}
