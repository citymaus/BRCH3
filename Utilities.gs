// Utilities.gs

function setActiveSpreadsheet(spreadsheetTabName) {   
  var spreadsheet = SpreadsheetApp.openById(Definitions.sheetId);
  var sheet = spreadsheet.getSheetByName(spreadsheetTabName);
  spreadsheet.setActiveSheet(sheet);
  SpreadsheetApp.getActiveSpreadsheet();

  return sheet;
}

function getByRangeName(spreadsheetTabName, rangeName) { 
  let row = null;
  let col = null;
  try {
    setActiveSpreadsheet(spreadsheetTabName);
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet();
    var range = activeSheet.getRangeByName(rangeName);

    row = range.getRow();
    col = range.getColumn();
  } catch (err) {
    let error = "Could not find range name: " + rangeName + ". Error: " + err;
    Logger.log(error);
  }
  return { row: row ?? rangeName, column: col ?? rangeName };
}

function formatDate(dateString) {
  return Utilities.formatDate(new Date(dateString), Definitions.timeZone, "MM/dd/yyyy");
} 

function formatCurrency(currency) {
  currency = currency.toString();
  let hasDollar = currency.indexOf("$") > -1;
  return (hasDollar ? "" : "$") + parseFloat(currency);  
}

function parseCurrency(currency) {
  return parseFloat(currency.replace("$", ""));
}

function calculateHabDues(totalDueCell) {
  let habDues = 0;
  let multipleAmountRegex = new RegExp(/(.*)(Amount: (.*) USD)(.*Quantity: (\d+))*.*\)/g);
  let descriptionGroup = 1;
  let amountGroup = 3;
  let quantityGroup = 5;
  var result = null;

  if (VERBOSE_LOGGING) {
    Logger.log("  > JOTFORM DUES (ignore camp dues here, takes tier from last date saved):")
    Logger.log("  " + totalDueCell.toString());
  }

  while((result = multipleAmountRegex.exec(totalDueCell)) !== null) {          
    let description = result[descriptionGroup];
    let amount = result[amountGroup];
    let quantity = result[quantityGroup];
    if (typeof result[quantityGroup] == 'undefined') {
      quantity = 1;
    }

    if (!description.toUpperCase().includes("CAMP DUES")) {
      habDues += (parseCurrency(amount) * parseInt(quantity));
    }
  }
  return habDues;
}

function calculateRequiredDues(paymentDate) {
  let formattedPaymentDate = new Date(paymentDate);
  let requiredDues = "999";

  for (let tier = 0; tier < DuesTiers.length; tier++) {
    let fromDate = new Date(DuesTiers[tier].fromDate + "/" + new Date().getFullYear());
    let toDate = new Date(DuesTiers[tier].toDate + "/" + new Date().getFullYear());

    if (formattedPaymentDate >= fromDate && formattedPaymentDate <= toDate) {
      requiredDues = DuesTiers[tier].amount;
      break;
    }
  }
  return parseFloat(requiredDues);
}

function findEarlierDate(firstDate, secondDate) {
  if (new Date(secondDate) < new Date(firstDate)) {
    return secondDate;
  }
  return firstDate;
}

function parseEmailBody(emailBody) {
  let parsedEmailBody = emailBody;
  let forwardedMessage = "---------- Forwarded message ---------";
  let forwardRegex = ".*" + forwardedMessage + "[\r\n]+(.*[\r\n]*)*";

  let foundMatch = emailBody.match(forwardRegex);

  if (foundMatch !== null) {
    parsedEmailBody = foundMatch[0].replace(forwardedMessage, "").trim();
  }
  return parsedEmailBody;
}

function getCamperNamesFromIdOverride(overrideId) {
  var tab = Definitions.paymentsTabName;
  var sheet = setActiveSpreadsheet(tab);  
  var manualIdCol = Columns.manualId - 1;
  var manualFirstCol = Columns.manualfirstName - 1;
  var manualLastCol = Columns.manualLastName - 1;
  var manualHashCol = Columns.manualHashName - 1;

  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var firstRow = Rows.manualIdOverrideRow;

  for (let i = firstRow; i < values.length; i++) {
    let paymentId = values[i][manualIdCol];
    if (paymentId != "" && Number.isInteger(paymentId)) {
      if (paymentId == overrideId) {  
        let firstName = values[i][manualFirstCol];
        let lastName = values[i][manualLastCol];
        let fullName = firstName + " " + lastName;
        let hashName = values[i][manualHashCol];
        return { firstName: firstName, lastName: lastName, fullName: fullName, hashName: hashName };
      }
    }
  }
  return null;  
}

function testNamedRanges() {
  setActiveSpreadsheet(Definitions.paymentsTabName);    
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let rowError = "";
  let colError = "";

  spreadsheet.toast("Validating expected named row ranges...");
  Object.keys(Rows).forEach(function(key,index) { 
    let result = Rows[key];
    if (!Number.isInteger(result)) {
      rowError = "ERROR: Could not find named range: " + result + ". Please check Data > Named Range for named ranges in your spreadsheet.";
      Logger.log(rowError);
    }
  });
  
  spreadsheet.toast("Validating expected named column ranges...");
  Object.keys(Columns).forEach(function(key,index) { 
    let result = Columns[key];    
    if (!Number.isInteger(result)) {
      colError = "ERROR: Could not find named range: " + result + ". Please check Data > Named Range for named ranges in your spreadsheet.";
      Logger.log(colError);
    }
  });

  if (rowError.length + colError.length == 0) {
    spreadsheet.toast("Success. All expected named rows and columns found.");
  } else {
    spreadsheet.toast("Error: " + rowError + " " + colError);
  }
}
