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

var Rows = {
  // Returns range row (integer) by named range
  // To name ranges, from spreadsheet: Data > Named Ranges
  firstPaymentsDataRow  : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.HeaderRow' ,'row'),
  paymentDueDataRow     : getByRangeName(Definitions.habOrdersTabName, 'HabOrders.TotalDue' ,'row'),
};

var Columns = {
  // Returns range column (integer) by named range
  // To name ranges, from spreadsheet: Data > Named Ranges
  paymentId         : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.PaymentId', 'col'),
  paymentIdOverride : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.PaymentIdOverride', 'col'),
  emailSubject      : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.EmailSubject', 'col'),
  emailMessage      : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.EmailMessage', 'col'),
  paymentAmount     : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.PaymentAmount', 'col'),
  paymentsTotal     : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.PaymentsTotal', 'col'),
  paymentDue        : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.PaymentDue', 'col'),
  paymentDate       : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.PaymentDate', 'col'),
  paymentDescription: getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.PaymentDescription', 'col'),
  balance           : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.Balance', 'col'),
  duesPaid          : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.DuesPaid', 'col'),
  hashName          : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.HashName', 'col'),
  firstName         : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.First', 'col'),
  lastName          : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.Last', 'col'),
  fullName          : getByRangeName(Definitions.paymentsTabName, 'ScrapedEmailData.FullName', 'col'),
  habOrderHashName  : getByRangeName(Definitions.habOrdersTabName, 'HabOrders.HashName', 'col'),
  habOrderTotalDue  : getByRangeName(Definitions.habOrdersTabName, 'HabOrders.TotalDue', 'col'),
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