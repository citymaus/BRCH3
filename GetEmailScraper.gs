// GetEmailScraper.gs
//-----------------------------------------------------------
// 2023-04-12 
// SKS (Whole Ass) 
// 
// Please ping me with questions or for debugging assistance!
// steph.stubler@gmail.com
// 
// For use with 2023 master camp spreadsheet:
// https://docs.google.com/spreadsheets/d/1wQxWVkbhc3m5-MdFntN-sf26Oy6rZHHnXnemCWqRVTo
//
// Link to public git repo for version tracking:
// https://github.com/citymaus/BRCH3
//-----------------------------------------------------------
function wholeAssEmailScraper(calledFromTimer = false){
  
  // Add script lock so trigger on a timer doesn't interfere with a manual scrape
  const lock = LockService.getScriptLock();
  if (!DEBUG){
    lock.waitLock(60 * 1000);
  }

  var startTime = Date.now();

  try {
    var paymentsSheet = setActiveSpreadsheet(Definitions.paymentsTabName);    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    if (!calledFromTimer) {
      spreadsheet.toast("Please wait. Scraping emails labeled [" + 
        Definitions.gmailLabel + "] for orders...");
    }

    const gmailLabel = GmailApp.getUserLabelByName(Definitions.gmailLabel);
    const emailThreads = gmailLabel.getThreads();

    let row = Rows.firstPaymentsDataRow + 1;
    let paymentId = 1;

    if (ADD_DATA_TO_SHEET) {
      // Clear rows
      let rowsToClear = 200;
      paymentsSheet.getRange(row, Columns.paymentId, row + rowsToClear, Columns.duesPaid).clearContent();
    }

    // Iterate from oldest email message first
    for (let i = emailThreads.length - 1; i >= 0; i--) 
    {
      var messages = emailThreads[i].getMessages();
      
      for (let j = messages.length - 1; j >= 0; j--) 
      {
        var emailSubject = emailThreads[i].getFirstMessageSubject();
        var emailDate = emailThreads[i].getLastMessageDate();
        var emailMsgOriginal = messages[j].getPlainBody();
        var emailMsg = parseEmailBody(emailMsgOriginal);
   
        var paymentData = new PaymentData(emailSubject, emailDate, emailMsg);
        var hasPartnerPayment = paymentData.camperNames.purchasePartnerIndex != null;

        if (hasPartnerPayment) {
          let partnerData = new PaymentData(emailSubject, emailDate, emailMsg, paymentData.camperNames.purchasePartnerIndex);
          
          let firstAmount = parseCurrency(paymentData.paymentDue);
          let secondAmount = parseCurrency(paymentData.paymentAmount) - firstAmount;

          paymentData.paymentAmount = formatCurrency(firstAmount);
          partnerData.paymentAmount = formatCurrency(secondAmount);

          // TODO, does not work if one or both partners makes another payment
          paymentData.paymentAmountTotal = formatCurrency(firstAmount);
          partnerData.paymentAmountTotal = formatCurrency(secondAmount);

          addDataToSheet(paymentsSheet,
            row, 
            paymentId, 
            emailSubject, 
            emailMsg, 
            partnerData);
            paymentId++;
            row++;   
        }

        addDataToSheet(paymentsSheet,
          row, 
          paymentId, 
          emailSubject, 
          emailMsg, 
          paymentData);

        paymentId++;
        row++;    
      }  
    }    
  } catch(err) {
    Logger.log("Exception caught: " + err.message);
  }

  var endTime = Date.now();
  var totalRuntime = ((endTime - startTime)/1000).toString();
  lock.releaseLock();

  Logger.log("Ran wholeAssEmailScraper in " + totalRuntime + " seconds.");
  Logger.log("----------------------------------------------------");
  Logger.log("SCRIPT COMPLETE");

  if (!calledFromTimer) {
    spreadsheet.toast("EMAIL SCRAPE SCRIPT COMPLETE. Runtime: " + totalRuntime + " seconds.");
  }
}

function manualIdOverride() {
  var tab = Definitions.paymentsTabName;
  var sheet = setActiveSpreadsheet(tab);    
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var paymentIdCol = Columns.paymentId - 1;
  var hashNameHeaderCol = Columns.hashName - 1;

  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var firstRow = Rows.paymentDueDataRow + 1;
  let count = 0;

  spreadsheet.toast("Finding \"?\" row entries and attempting to manually override...");
  
  for (let i = firstRow; i < values.length; i++) {
    let rowHasherName = values[i][hashNameHeaderCol];
    if (rowHasherName == "?") {
      let paymentId = values[i][paymentIdCol];
      
      let camperNames = getCamperNamesFromIdOverride(paymentId);
      Logger.log("Row " + i + " has blank hasher name.");
      if (camperNames != null) {
        Logger.log("Filling row " + i + " with " + camperNames.hashName);
        updateCamperNames(sheet, i + 1, camperNames);
        count++;
      }
    }
  }
  spreadsheet.toast("Finished manual overrides. Rows updated: " + count);
}

function addDataToSheet(paymentsSheet, row, paymentId, emailSubject, emailMsg, paymentData) {
    try {

      let balanceAfterFormula = parseCurrency(paymentData.paymentDue) - parseCurrency(paymentData.paymentAmountTotal);
      let duesPaid = !(balanceAfterFormula > 0);
      
      if (paymentData.camperNames.hashName == "?") {
        let camperNames = getCamperNamesFromIdOverride(paymentId);
        if (camperNames != null) {
          paymentData.camperNames.firstName = camperNames.firstName;
          paymentData.camperNames.lastName = camperNames.lastName;
          paymentData.camperNames.fullName = camperNames.firstName + " " + camperNames.lastName;
          paymentData.camperNames.hashName = camperNames.hashName;
        }
      }

      // Turn off for faster debugging
      if (ADD_DATA_TO_SHEET) {
        paymentsSheet.getRange(row, Columns.paymentId).setValue(paymentId);
        paymentsSheet.getRange(row, Columns.emailSubject).setValue(emailSubject);
        paymentsSheet.getRange(row, Columns.paymentDate).setValue(paymentData.paymentDate);     
        paymentsSheet.getRange(row, Columns.emailMessage).setValue(emailMsg);
        updateCamperNames(paymentsSheet, row, paymentData.camperNames);
        paymentsSheet.getRange(row, Columns.paymentAmount).setValue(paymentData.paymentAmount); 
        paymentsSheet.getRange(row, Columns.paymentsTotal).setValue(paymentData.paymentAmountTotal); 
        paymentsSheet.getRange(row, Columns.paymentDue).setValue(paymentData.paymentDue); 
        paymentsSheet.getRange(row, Columns.paymentDescription).setValue(paymentData.paymentDescription); 
        paymentsSheet.getRange(row, Columns.balance).setValue(balanceAfterFormula); 
        paymentsSheet.getRange(row, Columns.duesPaid).setValue(duesPaid ? "Yes" : ""); 
        
        Logger.log("ROW SUCCESSFULLY ADDED. (ID: " + paymentId + ", ROW: " + row
                    + ", Name: " + paymentData.camperNames.hashName + ")");
      }
    } catch (err) {
      Logger.log("Error writing payment to sheet: " + err.message);
    }
}

function updateCamperNames(paymentsSheet, row, camperNames) {
  if (ADD_DATA_TO_SHEET) {
    paymentsSheet.getRange(row, Columns.firstName).setValue(camperNames.firstName);
    paymentsSheet.getRange(row, Columns.lastName).setValue(camperNames.lastName);
    paymentsSheet.getRange(row, Columns.fullName).setValue(camperNames.fullName);
    paymentsSheet.getRange(row, Columns.hashName).setValue(camperNames.hashName);
  }
}