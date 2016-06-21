// This constant is written in column C for rows for which an email
// has been sent successfully.
var EMAIL_SENT = "Complete";


function sendEmails2() {
 
  //while loop 
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 1;  // First row of data to process
  
  
   var ss = SpreadsheetApp.getActiveSpreadsheet();
 var EMAIL_FAIL = ss.getSheetByName("EMAIL_FAIL");
 ss.setTabColor("A32929"); // Set the color to red.
EMAIL_FAIL.setTabColor(null);

  
  var data = sheet.getDataRange().getValues();//Checks for data in cells for the range
    
  for (var i = 1; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[0];  // First column
    var message = row[7];       // Second column
    var emailSent = row[2];     // Third column
    if (emailSent != EMAIL_SENT) {  // Prevents sending duplicates
      var subject = "Hotel confirmation";
      
      try {MailApp.sendEmail(emailAddress, subject, message);
      sheet.getRange(startRow + i, 3).setValue(EMAIL_SENT);
   // your script code here
          } catch(e) {sheet.getRange(startRow + i, 3).setValue(EMAIL_FAIL)
          
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getSheets()[0];
 // Returns the active cell
 var cell = sheet.getActiveCell();  
 var ss = SpreadsheetApp.getActiveSpreadsheet();
             
   // if the script code throws an error,
   // do something with the error here
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
  }}
