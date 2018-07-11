function LRS_TS_Send_Emails() {

 // Open sheet and tab
  var sheet = SpreadsheetApp.openById("1HrUhdbq2GtjkacJstKB3x46Ge0bqLJ0wQd_ay0rBqUM");
  var tab = sheet.getSheetByName("Sort By Audit Type");
 // Open Logs sheet and tabs
  var sheet = SpreadsheetApp.openById("1yN4KjuLfZVLrIkYbhqp4KTcNAh83zFm1bghjF3lmYtw");
  var logs = sheet.getSheetByName("Logs");


 // Assign variables
  var logsLastRow = logs.getLastRow() + 1;
  var logsStartCol = 1;
  var timeStamp = new Date();
  var EMAIL_SENT = "EMAIL_SENT";
  var startRow = 4; 
  var startCol = 1;
  var numRows = tab.getLastRow()-(startRow - 1);
  var numCol = tab.getLastColumn();
  var dataRange = tab.getRange(startRow, startCol, numRows, numCol);
  var data = dataRange.getValues();

 // Enter Email quota and Time Stamp in Logs tab
   logs.getRange(logsLastRow, logsStartCol).setValue(timeStamp);
   logs.getRange(logsLastRow, logsStartCol + 1).setValue(MailApp.getRemainingDailyQuota());
   logs.getRange(logsLastRow, logsStartCol + 2).setValue(MailApp.getRemainingDailyQuota());

 // Get array data
    for (var i = 0; i < data.length; ++i) {
    var row = data[i];

 // Create HTML String
    var msgHtml = "<b>K12 Learning Services Email Scorecard</b> "
    +"<br/>"
    +"<b>Agent Name: </b>" + row[3]
    +"<br/>"
    +"<b>Date of Interaction: </b>" + row[2]
    +"<br/>"
    +"<b>Audit Type: </b>" + row[7]
    +"<br/>"
    +"<b>Interaction Type: </b>" + row[5]
    +"<br/>"
    +"<b>Interaction ID: </b>" + row[0]
    +"<br/>"
    +"<b>Case Number: </b>" + row[4]
    +"<br/>"
    +"<b>Auto Fail: </b>" + row[24]
    +"<br/>"
    +"<hr />"
    +"<br/>"
    +"<b>Introduction and Documentation: </b>" + row[30]
    +"<br/>"
    +"<b>Customer Experience: </b>" + row[31]
    +"<br/>"
    +"<b>Closing: </b>" + row[32]
    +"<br/>"
    +"<b>Overall Percentage: </b>" + (row[34].toFixed(4))*100 + "%"
    +"<br/>"
    +"<hr />"
    +"<br/>"
    +"<b><i><u>INTRODUCTION AND DOCUMENTATION </b></i></u>"
    +"<br/>"
    +"<b>Greeting and Contact Details: </b>" + row[8]
    +"<br/>"
    +"<b>Notes: </b>" + row[9]
    +"<br/>"
    +"<br/>"
    +"<b>Documentation: </b>" + row[10]
    +"<br/>"
    +"<b>Notes: </b>" + row[11]
    +"<br/>"
    +"<hr />"
    +"<br/>"
    +"<b><i><u>CUSTOMER EXPERIENCE</b></i></u>"
    +"<br/>"
    +"<b>Listens and Responds Professionally: </b>" + row[12]
    +"<br/>"
    +"<b>Notes: </b>" + row[13]
    +"<br/>"
    +"<br/>"
    +"<b>Confidence and Competency: </b>" + row[14]
    +"<br/>"
    +"<b>Notes: </b>" + row[15]
    +"<br/>"
    +"<br/>"
    +"<b>Builds Relationship/Demonstrates Respect for the Customer: </b>" + row[16]
    +"<br/>"
    +"<b>Notes: </b>" + row[17]
    +"<br/>"
    +"<br/>"
    +"<b>Demonstrate Product, Platform, and Technical Knowledge: </b>" + row[18]
    +"<br/>"
    +"<b>Notes: </b>" + row[19]
    +"<br/>"
    +"<hr />"
    +"<br/>"
    +"<b><i><u>CLOSING</b></i></u>"
    +"<br/>"
    +"<b>Resolution, Completeness and Accuracy : </b>" + row[20]
    +"<br/>"
    +"<b>Notes: </b>" + row[21]
    +"<br/>"
    +"<br/>"
    +"<b>Case Conclusion: </b>" + row[22]
    +"<br/>"
    +"<b>Notes: </b>" + row[23]
    +"<br/>"
    +"<hr />"
    +"<br/>"
    +"<b><i><u>AUTO FAIL NOTES </b></b></u>"
    +"<br/>"
    + row[25]
    +"<br/>"
    +"<hr />"
    +"<br/>"
    +"<b><i><u>KUDOS </b></b></u>"
    +"<br/>"
    + row[26]
    +"<br/>"
    +"<hr />"
    +"<br/>"
    +"<b><i><u>COACHING NOTES </b></b></u>"
    +"<br/>"
    + row[27]
    + "<br/>"
    + "<hr />";
    
 // clear html tags and convert br to new lines for plain mail 
  var msgPlain = msgHtml.replace(/\<br\/\>/gi, '<br/>').replace(/(<([^>]+)>)/ig, "");
  var agentEmail = "james.robinson@pearson.com";
  var supEmail = "james.robinson@pearson.com";
//   var agentEmail = row[35];
//   var supEmail = row[37];
    var emailAddress = agentEmail + "," + supEmail + "," + "james.robinson@pearson.com";  
    var emailSent = row[38];    
    if (emailSent != EMAIL_SENT) {  // Prevents sending duplicates
    var subject = "QA score";
      MailApp.sendEmail(emailAddress, subject, msgPlain, ({htmlBody: msgHtml, noReply:true}));
      tab.getRange(startRow + i, numCol).setValue(EMAIL_SENT); // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
      
 // Log Agent names sent     
   Logger.log(row[3]);   
      
 // Pause script to allow Email quota to post
  Utilities.sleep(200);
      
//Enter new Email quota and Agent list on Logs tab 
logs.getRange(logsLastRow, logsStartCol + 2).setValue(MailApp.getRemainingDailyQuota());
logs.getRange(logsLastRow, logsStartCol + 3).setValue(Logger.getLog());  
    } 
  }
}