var dates_re = /[A-Z][a-z][a-z] \d\d\, 20\d\d/
var start_re = /\+ Details Hidden. Click to Show  /
var amt_re = /\(?\$[\d\,]*\.\d\d\)?/


function addTransactions() {
  //Takes user input from raw USAA data and outputs one line per transaction

  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Paste Unformatted');
  responseText = response.getResponseText();
  //responseText = "Date Description Account Category Amount + Details Hidden. Click to Show  Aug 03, 2018 HEB #425 AUSTIN TX   Signature Visa  Groceries      ($42.21) + Details Hidden. Click to Show  Aug 03, 2018 PAYPAL *PATREON MEMBE 402-935-7733 CA   USAA CASHBACK REWARDS CHECKING  Entertainment      ($1.00) + Details Hidden. Click to Show  Aug 02, 2018 AUSTIN 101 PROPE ACH ***********0668   USAA SAVINGS  Rent      ($1,400.00) + Details Hidden. Click to Show  Jul 31, 2018 AMAZON.COM AMZN.COM/BILL AMZN.COM/BILLWA   Signature Visa  Amazon      ($14.52)"
  Logger.log('%s', responseText);
 
  //trans = dates_re.exec(responseText);
  transactions = responseText.split(start_re);
 
  for (var i = 0; i < transactions.length; i++) {
    Logger.log("%s", transactions[i]);
    var temp = transactions[i]
    var date = dates_re.exec(temp);
    if (date) {
      temp = temp.split(date+' ')[1];
      temp = temp.split('  ');
      var desc = temp[0];
      var account = temp[1];
      var category = temp[2];
      var amount = temp[5];
     
      Logger.log([date[0], desc, account, category, amount])
      transactionRow = [date[0], amount, account, desc, category]
      
      //todo: add error handling for empty things
      
      var sheet = SpreadsheetApp.getActiveSheet();
      sheet.appendRow(transactionRow);
    }
    
    
  }
  
}  

function removeDuplicates() {
  //removes duplicate transactions (checks first 5 rows only)
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getDataRange().getValues();
  var data = sheet.getRange(1, 1, dataRange.length, 5).getValues();
  var newData = new Array();
  var actualData = new Array();
  for(i in data){
    var row = data[i];
    var duplicate = false;
    for(j in newData){
      if(row.join() == newData[j].join()){
        duplicate = true;
      }
    }
    if(!duplicate){
      
      newData.push(row);
      actualData.push(dataRange[i])
    }
  }
  sheet.clearContents();
  sheet.getRange(1, 1, actualData.length, actualData[0].length)
      .setValues(actualData);
}  
  
   
  

function hilightDuplicates() {
  //highlights duplicates (checks first 5 columns only)
  var sheet = SpreadsheetApp.getActiveSheet();
  var numCols = sheet.getLastColumn();
  
  var dataRange = sheet.getDataRange().getValues();
  var data = sheet.getRange(1, 1, dataRange.length, 5).getValues();

  
  for(i in data){
    for(j in data){
      if(i != j){
        if(data[i].join() == data[j].join()){
          first =+i;
          second =+j;
          sheet.getRange((first+1),1, 1, numCols).setBackground("yellow");
        }
      }
    }
  }

}  



function onOpen() {
  // Add a custom menu to the spreadsheet.
  SpreadsheetApp.getUi() // Or DocumentApp, SlidesApp, or FormApp.
      .createMenu('Custom')
      .addItem('Add Transaction(s)', 'addTransactions')
      .addItem('Highlight Duplicates', 'hilightDuplicates')
      .addItem('Remove Duplicates', 'removeDuplicates')
      .addToUi();
}
