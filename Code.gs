// Runs when editor opens spreadsheet
// Add trigger with From spreadsheet - On open
function onOpen(source) {
  Logger.log("onOpen called. Trigger source: " + Object.values(source));

  var sheet = setActiveSpreadsheet(Definitions.paymentsTabName);
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  spreadsheet.toast("Adding 'Email Scraper' to menu...");
  var menuEntries = [ 
    {name: "Verify named ranges exist", functionName: "testNamedRanges"},
    {name: "Load order emails...", functionName: "wholeAssEmailScraper"} 
  ];
  spreadsheet.addMenu("Email Scraper", menuEntries);
}

// Runs when Time-Driven event source trigger is called.
function onTimer(source) {
  Logger.log("onTimer called. Trigger source: " + Object.values(source));
  
  // Original code:
  //getEmails();
  
  // 2023-04-11:
  wholeAssEmailScraper(true);
}

//original script: https://stackoverflow.com/questions/75774623/gmailapp-not-returning-all-messages-in-thread
function getEmails(){
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Payments');
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const label = GmailApp.getUserLabelByName("taxes/payments");
  const threads = label.getThreads();

  let row = 17;

  for (let i = 0; i < threads.length; i++) {

    const messages = threads[i].getMessages();
    
    for (var j = 0; j < messages.length; j++) {

      const msg = messages[j].getPlainBody(); 

      sheet.getRange(row,1).setValue(threads[i].getFirstMessageSubject());
      sheet.getRange(row,2).setValue(threads[i].getLastMessageDate());     
      sheet.getRange(row,3).setValue(msg);

      row++;
    
    }  
  }
}


