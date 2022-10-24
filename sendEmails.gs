// This constant is written in the last column for rows for which an email has been sent successfully.
var EMAIL_SENT = 'EMAIL_SENT';

/**
 * Sends non-duplicate emails with data from the current spreadsheet.
 */
function sendEmails() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Email Runway'); 

  var startRow = 2; // First row of data to process
  var range = sheet.getDataRange().getValues()
  var numRows = range.length; // Number of rows to process

  // Fetch the range of cells A2:E
  var dataRange = sheet.getRange(startRow, 1, numRows, 5);
  var data = dataRange.getValues();

  var scooter = DriveApp.getFileById('1Jq72_R67uK_viei6VB4uHy1ZHfdNd9Mk');
  var top = DriveApp.getFileById('1b3UjaWn7ZRlxSYen2D_F1OCiRPZtw6Wa');
  var bottom = DriveApp.getFileById('1p-WBSxKvdWxoVw0E2SQlQslL7L0EJ6oX');
  var twitter = DriveApp.getFileById('16SWWikw6MUAfMzvhxImCL2lF2S9NqrnI');
  var instagram = DriveApp.getFileById('1wnXPWkzKS314Cpez5XCug-ZHN8UZXZOW');
  var facebook = DriveApp.getFileById('11Lqv_NzRe1BdrFKzaiEt6B1AwY5vy03K');
  var googleplay = DriveApp.getFileById('1uDJPVqTxj7udJ76cO9Flz58Gu7uGrn8z');
  var applestore = DriveApp.getFileById('1PgJUWuO3Tti5MXprNYKKm8HxwkmeGG6s');

  for (var i = 0; i < numRows; ++i) {
     var row = data[i];
     var emailAddress = row[0]; // First column
     var discountCode = row[1]; // Second column
     var header = row[2]; // Third column
     var textBody = row[3]; // Fourth column
     var emailSent = row[4]; // Fifth column
     if (emailSent !== EMAIL_SENT && emailAddress !== "") { // Prevents sending duplicates and prevents trying to send email to blank cells
      var subject = 'LINK PROMO';
      var html = HtmlService.createTemplateFromFile('EmailTemplate');
      html.discountCode = discountCode;
      html.header = header;
      html.textBody = textBody;

      GmailApp.sendEmail(emailAddress,subject,discountCode,
        {
          htmlBody:html.evaluate().getContent(),
          inlineImages:
          {
            scooterPic: scooter.getBlob(), 
            topLogo: top.getBlob(), 
            bottomLogo: bottom.getBlob(), 
            twitterLogo: twitter.getBlob(),
            instagramLogo: instagram.getBlob(),
            facebookLogo: facebook.getBlob(),
            googleplayLogo: googleplay.getBlob(),
            applestoreLogo: applestore.getBlob()  
          }
        }
        );

      sheet.getRange(startRow + i, 5).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
   }
  }
}
