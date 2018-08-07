//var dates_re = new RegExp("[A-Z][a-z][a-z] \d\d\, 20\d\d");
//var dates_re = /[a-z]+/


function getMonday(d_str) {
  
  d_parts = d_str.split(/\W/)
  d = new Date(Date.parse(d_str));
  Logger.log(d)
  var day = d.getDay(),
      diff = d.getDate() - day + (day == 0 ? -6:1); // adjust when day is sunday
  return new Date(d.setDate(diff));
}


function addTransactions() {
  var dates_re = /[A-Z][a-z][a-z] \d\d\, 20\d\d/
  var start_re = /\+ Details Hidden. Click to Show  /
  var amt_re = /\(?\$[\d\,]*\.\d\d\)?/
  var check = /        View Check for .* \(Opens a pop-up layer\)\. /
  
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Paste Unformatted from Money Managwer');
  responseText = response.getResponseText();
  //responseText = "Date Description Account Category Amount + Details Hidden. Click to Show  Aug 09, 2018 HEB #425 AUSTIN TX   Signature Visa  Groceries      ($42.21) + Details Hidden. Click to Show  Aug 03, 2018 PAYPAL *PATREON MEMBE 402-935-7733 CA   USAA CASHBACK REWARDS CHECKING  Entertainment      ($1.00) + Details Hidden. Click to Show  Aug 02, 2018 AUSTIN 101 PROPE ACH ***********0668   USAA SAVINGS  Rent      ($1,400.00) + Details Hidden. Click to Show  Jul 31, 2018 AMAZON.COM AMZN.COM/BILL AMZN.COM/BILLWA   Signature Visa  Amazon      ($14.52)"
  //responseText = "Date Description Account Category Amount + Details Hidden. Click to Show  Aug 02, 2018 CAVK Tips   REGULAR SAVINGS  CAVK      $368.25 + Details Hidden. Click to Show  Jul 31, 2018 INTEREST PAID   USAA SAVINGS  Interest      $0.48 + Details Hidden. Click to Show  Jul 12, 2018 House sitting        View Check for House sitting (Opens a pop-up layer). USAA CASHBACK REWARDS CHECKING  Other Income      $175.00 + Details Hidden. Click to Show  Jul 12, 2018 CAVK        View Check for CAVK (Opens a pop-up layer). USAA CASHBACK REWARDS CHECKING  CAVK      $186.64 + Details Hidden. Click to Show  Jul 02, 2018 CAVK        View Check for CAVK (Opens a pop-up layer). USAA CASHBACK REWARDS CHECKING  CAVK      $135.59 + Details Hidden. Click to Show  Jun 29, 2018 INTEREST PAID   USAA SAVINGS  Interest      $0.34 + Details Hidden. Click to Show  Jun 25, 2018 CAVK        View Check for CAVK (Opens a pop-up layer). USAA CASHBACK REWARDS CHECKING  CAVK      $127.55 + Details Hidden. Click to Show  Jun 21, 2018 CAVK        View Check for CAVK (Opens a pop-up layer). USAA CASHBACK REWARDS CHECKING  CAVK      $236.78 + Details Hidden. Click to Show  Jun 12, 2018 PAYPAL *RED CROSS   Chase Freedom  CAVK      ($25.00) + Details Hidden. Click to Show  Jun 11, 2018 OR REVENUE DEPT ORSTTAXRFD ***********9122   USAA SAVINGS  Other Income      $182.30 + Details Hidden. Click to Show  Jun 11, 2018 CAVK        View Check for CAVK (Opens a pop-up layer). USAA CASHBACK REWARDS CHECKING  CAVK      $197.13 + Details Hidden. Click to Show  May 31, 2018 INTEREST PAID   USAA SAVINGS  Interest      $0.46 + Details Hidden. Click to Show  May 31, 2018 UNIV TX AUSTIN PAYROLL ***********2786   USAA CASHBACK REWARDS CHECKING  UT - TA      $1,391.81 + Details Hidden. Click to Show  May 14, 2018 CAVK        View Check for CAVK (Opens a pop-up layer). USAA CASHBACK REWARDS CHECKING  CAVK      $183.76 + Details Hidden. Click to Show  May 07, 2018 CAVK        View Check for CAVK (Opens a pop-up layer). USAA CASHBACK REWARDS CHECKING  CAVK      $42.02 + Details Hidden. Click to Show  May 07, 2018 CAVK        View Check for CAVK (Opens a pop-up layer). USAA CASHBACK REWARDS CHECKING  CAVK      $132.91 + Details Hidden. Click to Show  Apr 30, 2018 INTEREST PAID   USAA SAVINGS  Interest      $0.51 + Details Hidden. Click to Show  Apr 30, 2018 UNIV TX AUSTIN PAYROLL ***********2786   USAA CASHBACK REWARDS CHECKING  UT - TA      $1,408.26 + Details Hidden. Click to Show  Apr 10, 2018 VENMO VERIFYBAN ***********4115   USAA CASHBACK REWARDS CHECKING  Other Income      $0.20 + Details Hidden. Click to Show  Apr 10, 2018 VENMO VERIFYBAN ***********4118   USAA CASHBACK REWARDS CHECKING  Other Income      $0.31 + Details Hidden. Click to Show  Apr 06, 2018 IRS TREAS 310 TAX REF ***********0918   USAA SAVINGS  Refunds/Adjustments      $822.00 + Details Hidden. Click to Show  Mar 30, 2018 INTEREST PAID   USAA SAVINGS  Interest      $0.39 + Details Hidden. Click to Show  Mar 30, 2018 UNIV TX AUSTIN PAYROLL ***********2786   USAA CASHBACK REWARDS CHECKING  UT - TA      $1,407.38 + Details Hidden. Click to Show  Mar 22, 2018 CASH REWARDS CREDIT   USAA SAVINGS  Rewards      $5.00 + Details Hidden. Click to Show  Feb 28, 2018 INTEREST PAID   USAA SAVINGS  Interest      $0.39"
  Logger.log('%s', responseText);
 
  //trans = dates_re.exec(responseText);
  transactions = responseText.split(start_re);
 
  for (var i = 0; i < transactions.length; i++) {
    Logger.log("raw: %s", transactions[i]);
    var temp = transactions[i].replace(check, '   ')
    var date_str = dates_re.exec(temp);
    if (date_str) {
      
      
      var date = new Date(Date.parse(date_str[0]));
      var monday = getMonday(date_str[0]);
      
      
      temp = temp.split(date_str+' ')[1];
      temp = temp.split('  ');
      Logger.log('TEMP: %s', temp)
      var desc = temp[0];
      var account = temp[1].trim();
      var category = temp[2];
      var amount = temp[5];
      
       
      
      //Logger.log([date, desc, account, category, amount, monday])
      transactionRow = [date, amount, account, desc, category, monday]
      
      //todo: add error handling for empty things
      
      var sheet = SpreadsheetApp.getActiveSheet();
      sheet.appendRow(transactionRow);
    }
    
    
  }
  
}  



function removeDuplicates() {
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
      .addItem('Add Venmo Transaction(s)', 'addTransactions')
      .addItem('Highlight Duplicates', 'hilightDuplicates')
      .addItem('Remove Duplicates', 'removeDuplicates')
      .addToUi();
}
