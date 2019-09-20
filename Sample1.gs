// Sort Timestamp column from newest to oldest: 

function Sort_Timestamp(event){
  var sheet2 = SpreadsheetApp.getActiveSheet();
  var editedCell = sheet.getActiveCell();

  var columnToSortBy = 5;                          //Column 5 is D; the timestamp column
  var tableRange = "A:P";                         //include range of sheet

  if(editedCell.getColumn() == columnToSortBy){   
    var range = sheet.getRange(tableRange);
    range.sort( { column : columnToSortBy, ascending: false } );
  }
}

// Automated Emails:

var ss = SpreadsheetApp.getActive();
var sheet = SpreadsheetApp.getActiveSheet();

function EmailAutomation(e) {                      
 var row = e.range.getRow();
 var col = e.range.getColumn();
 var status = e.range.getValue();
 var email = sheet.getRange(row,6).getValue();
 var ir = sheet.getRange(row,2).getValue();
 var emailSent_InProgress = sheet.getRange(row,12).getValue();       
 var emailSent_Completed = sheet.getRange(row,13).getValue(); 
 if (col == 3 && status=="In Progress" && ir!="") {        
    if(emailSent_InProgress == "") {                              
      sendEmail_InProgress(email);    
      sheet.getRange(row,12).setValue("Email has been sent; status 'Request Acknowledged'");
      sheet.getRange(row,14).setValue(new Date()).setNumberFormat('M/D/YYYY hh:mm:ss'); 
      SpreadsheetApp.flush();
    }
     } 
    else if (col == 3 && status=="Completed" && ir!="") {
    if(emailSent_Completed == "") {
    sendEmail_Completed(email);
    sheet.getRange(row,13).setValue("Email has been sent; status 'Completed'");
    sheet.getRange(row,15).setValue(new Date()).setNumberFormat('M/D/YYYY hh:mm:ss');   
    SpreadsheetApp.flush();
}
  }
  else if (col == 3 && status== "Escalated") {        //if at any point the current status is set to "Escalated to L3", this will change cell in column P to "Yes"
  sheet.getRange(row,16).setValue("Yes");
  }
  
  
  function sendEmail_InProgress(email) {
  var subject = ("Request Acknowledged, IR#: " + ir);
  var html = HtmlService.createHtmlOutputFromFile("In_Progress_Email").getContent();
     MailApp.sendEmail({                                          
     to: email,                                                    //recipient                
     subject: subject,                                             //subject
     htmlBody: html,                                               //body
     noReply: true                                                 //email will be from noreply
     });
  }
  
  
  function sendEmail_Completed(email) {
  var subject = ("Request Completed, IR#: " + ir);
  var html = HtmlService.createHtmlOutputFromFile("Completed_Email").getContent();
     MailApp.sendEmail({                                           
     to: email,                                                    //recipient                
     subject: subject,                                             //subject
     htmlBody: html,                                               //body
     noReply: true                                                 //email will be from noreply
     });
  }

}
