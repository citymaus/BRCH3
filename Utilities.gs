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

function calculateRequiredDues(paymentDate) {
  let formattedPaymentDate = new Date(paymentDate);
  let requiredDues = "999";

  for (let tier = 0; tier < DuesTiers.length; tier++) {
    let fromDate = new Date(DuesTiers[tier].from + "/" + new Date().getFullYear());
    let toDate = new Date(DuesTiers[tier].to + "/" + new Date().getFullYear());

    if (formattedPaymentDate >= fromDate && formattedPaymentDate <= toDate) {
      requiredDues = DuesTiers[tier].amount;
      break;
    }
  }
  return parseFloat(requiredDues);
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
