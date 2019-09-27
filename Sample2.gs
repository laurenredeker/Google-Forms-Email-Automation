// this function organizes the rows from 'Form Responses' sheet and pastes them into 'Paste Values' sheet
// Giving credit to those who helped me with this on Stackoverflow: https://stackoverflow.com/questions/55716140/google-apps-script-sheets-forms-data-manipulation-deleting-rows-if-certain-cel/55756156?noredirect=1#comment98227752_55756156

function organize_my_rows(event) {

  // setup this function as an installable trigger OnFormSubmit

  // set up spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourcename = "Form Responses";
  var targetname = "Paste Values";
  var sourcesheet = ss.getSheetByName(sourcename);
  var targetsheet = ss.getSheetByName(targetname);

  // get the response row and range
  var sourcerow = event.range.getRow();
  //Logger.log("DEBUG: Response row = "+sourcerow); //DEBUG

  // range is from Column A to Column K: 11 columns; 2=common; 5 name "section" questions PLUS "do you want to tell me again?"
  var sourcerange = sourcesheet.getRange(sourcerow, 1, 1, 11);  //(source row, column #, number of rows, number of columns);
  //Logger.log("DEBUG: Source range = "+sourcerange.getA1Notation()); //DEBUG

  // get the response data
  var sourcedata = sourcerange.getValues();

  // find the number of rows already populated on the target
  var Bvals = targetsheet.getRange("B1:B").getValues();
  var Blast = Bvals.filter(String).length;
  //Logger.log("DEBUG: Blast = "+Blast); //DEBUG

  // establish some variables
  var datastart = 2; // the first two columns are common data (we get the timestamp + person's email address just once)
  var dataqty = 2; // the first 4 responses have 2 columns of response data
  var printcount = 0; // status counter
  var responsecount = 0; // column status counter
  var responsedata = []; // array to compile responses

  // process the first section
  if (printcount == 0) {
    responsedata = [];

    // get the timestamp and email
    responsedata.push(sourcedata[0][0]);
    responsedata.push(sourcedata[0][1]);

    //get the responses for the next 2 questions
    for (i = datastart; i < (datastart + dataqty); i++) {
      responsedata.push(sourcedata[0][i]);
    }

    // define the target range
    // the last line (Blast)plus one line plus the print count; column B; 1 row; 4 columns
    var targetrange = targetsheet.getRange(Blast + 1 + printcount, 1, 1, 4); // getRange(source row, column #, number of rows, number of columns);
    // paste the values to the target
    targetrange.setValues([responsedata]);

    // update variables
    responsecount = i; // copy the value of i
    printcount++; // update status counter
    responsedata = []; // clear the array ready for the next section
  }
  // end opening response

  // build routine for 2nd to 4th sections
  for (z = 2; z < 5; z++) {

    //Make sure not to double count the first section
    if (printcount > 0 && printcount < 5) {

      // test if the next section exists
      if (sourcedata[0][i - 1] == "Yes") {
        // Yes for next section
        //Logger.log("DEBUG: value = "+sourcedata[0][i-1]);  //DEBUG

        // get the timestamp and email
        responsedata.push(sourcedata[0][0]);
        responsedata.push(sourcedata[0][1]);

        //get the responses for the next 10 questions
        for (i = responsecount; i < (responsecount + dataqty); i++) {
          responsedata.push(sourcedata[0][i]);
          //Logger.log("DEBUG: data: "+sourcedata[0][i]);//DEBUG
        }

        // define the target range
        // the last line (Blast) plus one line plus the print count; column B; 1 row; 4 columns
        targetrange = targetsheet.getRange(Blast + 1 + printcount, 1, 1, 4);
        // paste the values to the target
        targetrange.setValues([responsedata]);

        // update variables
        responsecount = i;
        printcount++;
        responsedata = [];
      } else {
        // NO for next section
      }
      // end routine if the next section exists
    } // end routine for the next section
  } // end routine for sections 2, 3 and 4

  // checking for 5th response
  if (printcount == 4) {

    // test if response exists
    if (sourcedata[0][i - 1] == "Yes") {
      // Yes for next section
      //Logger.log("DEBUG: value = "+sourcedata[0][i-1]);  //DEBUG

      // get the timestamp and email
      responsedata.push(sourcedata[0][0]);
      responsedata.push(sourcedata[0][1]);

      //get the responses for the next (1) question
      for (i = responsecount; i < (responsecount + dataqty - 1); i++) {
        responsedata.push(sourcedata[0][i]);
        //Logger.log("DEBUG: data: "+sourcedata[0][i]);//DEBUG
      }

      // define the target range
      // the last line (Blast) plus one line plus the print count; column B; 1 row; 3 columns only (because 5th round doesn't ask 'do you want to tell me again')
      targetrange = targetsheet.getRange(Blast + 1 + printcount, 1, 1, 3);
      // paste the values to the target
      targetrange.setValues([responsedata]);
    } else {
      // NO for next section
    }
  }
  // end routine for the 5th section

}
