// Definitions.gs

var DEBUG = false;              // When true, populates test spreadsheet. Ignores script locking that waits for triggered runs
var DEBUG_EMAIL = false;        // When true, loads emails from debug account / label
var VERBOSE_LOGGING = true;     // When false, can speed up script run. (Max time limit is 6 mins)
var ADD_DATA_TO_SHEET = true;   // Set to false to speed up debugging individual records, does not add data to spreadsheet

var Definitions = {
  timeZone: "PDT",
  sheetId: getSheetProperties().sheetId,
  gmailLabel: getSheetProperties().gmailLabel,
  paymentsTabName: "Payments",
  formResponsesTabName: "Form Responses",
  habOrdersTabName: "Hab Orders"
};

var PaymentSource = {
  Zelle: "ZELLE",
  GPay: "GPAY",
  BRCH3Website: "BRCH3Website"
};

var DuesTiers = [
  { 
    from: "3/8",
    to: "3/31",
    amount: "235"
  },
  { 
    from: "4/1", 
    to: "5/31", 
    amount: "255"
  },
  { 
    from: "6/1",
    to: "7/7", 
    amount: "275"
  },
  { 
    from: "7/8", 
    to: "8/4", 
    amount: "290"
  }
]

var Rows = {
  // Returns range row (integer) by named range
  // To name ranges, from spreadsheet: Data > Named Ranges
  firstPaymentsDataRow  : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.HeaderRow').row,
  paymentDueDataRow     : getByRangeName(Definitions.habOrdersTabName, 'HabOrders.TotalDue').row,
};

var Columns = {
  // Returns range column (integer) by named range
  // To name ranges, from spreadsheet: Data > Named Ranges
  paymentId         : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.PaymentId').column,
  paymentIdOverride : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.PaymentIdOverride').column,
  emailSubject      : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.EmailSubject').column,
  emailMessage      : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.EmailMessage').column,
  paymentAmount     : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.PaymentAmount').column,
  paymentsTotal     : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.PaymentsTotal').column,
  paymentDue        : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.PaymentDue').column,
  paymentDate       : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.PaymentDate').column,
  paymentDescription: getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.PaymentDescription').column,
  balance           : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.Balance').column,
  duesPaid          : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.DuesPaid').column,
  hashName          : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.HashName').column,
  firstName         : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.First').column,
  lastName          : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.Last').column,
  fullName          : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.FullName').column,
  habOrderHashName  : getByRangeName(Definitions.habOrdersTabName, 'HabOrders.HashName').column,
  habOrderTotalDue  : getByRangeName(Definitions.habOrdersTabName, 'HabOrders.TotalDue').column,
};

function getSheetProperties() {  
  var sheetId = "1wQxWVkbhc3m5-MdFntN-sf26Oy6rZHHnXnemCWqRVTo";
  var gmailLabel = "taxes/payments";

  if (DEBUG) {  
    var sheetId = "1Fgouf6lhCP70HZFbuja_yTXX0byQipMqMvPom9NJgBs";
  }
  if (DEBUG_EMAIL) {
    var gmailLabel = "Burning Man/BRCH3";
  }
  return { sheetId: sheetId, gmailLabel: gmailLabel }
}