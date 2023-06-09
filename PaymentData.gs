// PaymentData.gs
//-----------------------------------------------------------
// 2023-04-12 
// SKS (Whole Ass) 
// 
// Please ping me with questions or for debugging assistance!
// steph.stubler@gmail.com
//
// This is where the meat of parsing emails with regexes lives.
//-----------------------------------------------------------
class PaymentData {
  constructor(emailSubject, emailDate, emailContent, partnerIndex = null) {
    this.emailSubject = emailSubject;
    this.emailDate = emailDate;
    this.emailContent = emailContent;

    if (VERBOSE_LOGGING) {
      Logger.log("Email subject: " + this.emailSubject); 
      Logger.log("Email date: " + this.emailDate); 
      Logger.log("Email content: " + this.emailContent); 
    }

    // Items to parse:
    this.paymentSource = this.getPaymentSource(); 
    this.paymentDescription = this.getPaymentDescription();
    this.paymentAmount = this.getPaymentAmount(); 
    this.paymentDate = this.getPaymentDate(); 
    this.partnerRow = partnerIndex;
    this.camperNames = this.getCamperData(this.getCamperName(), this.paymentDescription, this.partnerRow);

    let postResult = this.makeMultipleRowCalculations();
    this.paymentDue = this.getPaymentDue(postResult.earliestPaymentDate);
    this.paymentAmountTotal = postResult.totalPaid;

    if (VERBOSE_LOGGING) {
      Logger.log("Payment source: " + this.paymentSource); 
      Logger.log("Payment date: " + this.paymentDate);
      Logger.log("Payment description: " + this.paymentDescription);
      Logger.log("Payment amount: " + this.paymentAmount);
      Logger.log("Payments total made: " + this.paymentAmountTotal);
      Logger.log("Payment due: " + this.paymentDue);
      Logger.log("Camper name: " + this.camperNames.fullName);
      Logger.log("Hash name: " + this.camperNames.hashName);
      Logger.log("  PAYMENT DATA RETRIEVAL ROW: DONE");
    }
  }  

 getCamperData(moneySender, paymentDescription, partnerIndex = null) {
      let parsedFirstName = "?";
      let parsedLastName = "?";
      let parsedFullName = "?";
      let parsedHashName = "?";
      let foundData = false;
      let foundHashNameCount = 0;
      
      var tab = Definitions.formResponsesTabName;
      var formResponseSheet = setActiveSpreadsheet(tab);

      var hashNameHeaderCol = getByRangeName(tab, 'FormResponses.HashName').column - 1;
      var firstNameHeaderCol = getByRangeName(tab, 'FormResponses.FirstName').column - 1;
      var lastNameHeaderCol = getByRangeName(tab, 'FormResponses.LastName').column - 1;

      var dataRange = formResponseSheet.getDataRange();
      var values = dataRange.getValues();

      //--------------------------------------------------------------------
      // Use partner index if defined (already passed through this function)      
      //--------------------------------------------------------------------
      if (partnerIndex != null) {        
        parsedFirstName = values[partnerIndex][firstNameHeaderCol];
        parsedLastName = values[partnerIndex][lastNameHeaderCol];
        parsedHashName = values[partnerIndex][hashNameHeaderCol];
        parsedFullName = parsedFirstName + " " + parsedLastName;
        partnerIndex = null;
      } else {
 
        //--------------------------------------------------------------------
        // Start from possible hash name inside payment description   
        //--------------------------------------------------------------------
        var purchasePartnerIndex = null;
        for (var i = 1; i < values.length; i++) {
          let hashName = values[i][hashNameHeaderCol];

          var hasHashNameToFind = paymentDescription.length > 0 && hashName.length > 0;
          let foundHashName = hasHashNameToFind ? paymentDescription.toUpperCase().includes(hashName.toUpperCase())
                                                : false;   
          let foundComplexHashName = hasHashNameToFind ? this.tryComplexCamperName(hashName, paymentDescription)
                                                : false;   

          if (foundHashName) {
            Logger.log(" > Found camper by hash name: " + hashName);
            foundData = true;
            foundHashNameCount++;
            if (foundHashNameCount == 1) {
              parsedFirstName = values[i][firstNameHeaderCol];
              parsedLastName = values[i][lastNameHeaderCol];
              parsedHashName = values[i][hashNameHeaderCol];
              parsedFullName = parsedFirstName + " " + parsedLastName;
            } else {
              purchasePartnerIndex = i;
              Logger.log(" > Purchase was also made for partner: " + hashName);
            }
          } else {
            if (foundComplexHashName.length > 0) {
              let diff = hashName.length - foundComplexHashName.length;
              let acceptableDiff = 12; // Very scientific
              let similarToExpectedHashName = diff < acceptableDiff;
              if (similarToExpectedHashName) {
                Logger.log(" > Found camper by complex hash name: " + foundComplexHashName);
                foundData = true;
                foundHashNameCount++;
                if (foundHashNameCount == 1) {
                  parsedFirstName = values[i][firstNameHeaderCol];
                  parsedLastName = values[i][lastNameHeaderCol];
                  parsedHashName = values[i][hashNameHeaderCol];
                  parsedFullName = parsedFirstName + " " + parsedLastName;
                } else {
                  purchasePartnerIndex = i;
                  Logger.log(" > Purchase was also made for partner: " + hashName);
                }
              }
            }
          }       
        }

        //--------------------------------------------------------------------
        // Not found by hash name, try first AND last name
        //--------------------------------------------------------------------
        if (!foundData) { 
          let hasFirstAndLastNameInMoneySender = moneySender.includes(" ") && moneySender.length > 1;
          let hasNameInMoneySender = moneySender.length > 1;
          let foundMultipleCandidates = false;
          let candidate = null;

          for (var i = 1; i < values.length; i++) {
            let firstName = values[i][firstNameHeaderCol].length > 0 ? values[i][firstNameHeaderCol] : "unknown";
            let lastName  = values[i][lastNameHeaderCol].length > 0 ? values[i][lastNameHeaderCol] : "unknown";
            
            let foundLastNameInDescription = paymentDescription.toUpperCase().includes(lastName.toUpperCase());
            let foundFirstNameInDescription = paymentDescription.toUpperCase().includes(firstName.toUpperCase());
            
            let foundLastNameInMoneySender = moneySender.toUpperCase().includes(lastName.toUpperCase());
            let foundFirstNameInMoneySender = moneySender.toUpperCase().includes(firstName.toUpperCase());

            if ((foundLastNameInDescription && foundFirstNameInDescription)
              || (hasFirstAndLastNameInMoneySender 
                  && (foundLastNameInMoneySender && foundFirstNameInMoneySender))) {
              Logger.log(" > Found camper by first and last name. First: " + firstName + " Last: " + lastName);
              foundData = true;
              parsedFirstName = values[i][firstNameHeaderCol];
              parsedLastName = values[i][lastNameHeaderCol];
              parsedHashName = values[i][hashNameHeaderCol];
              parsedFullName = parsedFirstName + " " + parsedLastName;
              candidate = i;
              break;
            }
            //--------------------------------------------------------------------
            // Try with first OR last name only
            //--------------------------------------------------------------------
            else if (hasNameInMoneySender && 
                      (foundFirstNameInMoneySender || foundLastNameInMoneySender)
                    ) {
              if (candidate == null) {
                candidate = i;
              }
              else {
                foundMultipleCandidates = true;
                break;
              }
            }
          }

          
          //--------------------------------------------------------------------
          // Do not accept as match if multiple candidates are found. Otherwise,
          // assume this is a camper match.
          //--------------------------------------------------------------------
          if (!foundMultipleCandidates && candidate != null) {
            Logger.log(" > Found camper by (unique) first or last name.");
            Logger.log("  > First name: " + values[candidate][firstNameHeaderCol]);
            Logger.log("  > Last name: " + values[candidate][lastNameHeaderCol]);
            foundData = true;
            parsedFirstName = values[candidate][firstNameHeaderCol];
            parsedLastName = values[candidate][lastNameHeaderCol];
            parsedHashName = values[candidate][hashNameHeaderCol];
            parsedFullName = parsedFirstName + " " + parsedLastName;
          }
        }

        if (!foundData) {
            
            //--------------------------------------------------------------------
            // No match found, default to full name from money sender parse
            //--------------------------------------------------------------------
            parsedFullName = moneySender;
            Logger.log(" > ERROR: Could not find camper by hash, first, or last name.");
            Logger.log("  > Money sender: " + moneySender);
            Logger.log("  > Payment description: " + paymentDescription);
        }
      }

      // Return to current Payments sheet      
      setActiveSpreadsheet(Definitions.paymentsTabName);

      return { 
        firstName: parsedFirstName, 
        lastName: parsedLastName, 
        fullName: parsedFullName, 
        hashName: parsedHashName, 
        purchasePartnerIndex: purchasePartnerIndex };
  }

  //--------------------------------------------------------------------
  // Email parsing with regex section!
  //--------------------------------------------------------------------
  getPaymentSource() {
    if (this.emailContent.toUpperCase().includes("ZELLE")) {
      return PaymentSource.Zelle;
    }
    if (this.emailContent.toUpperCase().includes("A WANKER GAVE US MONEY")) {
      return PaymentSource.BRCH3Website;
    }
    if (this.emailContent.toUpperCase().includes("PAYPAL")) {
      return PaymentSource.PayPal;
    }
    return PaymentSource.GPay;
  }

  getPaymentDescription() {
    let content = this.emailContent;
    let description = "N/A";

    switch (this.paymentSource) {
      case (PaymentSource.Zelle): {
        var regex = "Description *.(.*)";
        var group = 1; // Item in parentheses
        try {
          description = content.match(regex)[group];
        } catch {
          // Description not found.
        }
        break;
      }
      case (PaymentSource.GPay): {
        var regex = ".*[\u201c|\u201C](.*)[\u201d|\u201D]"; // \u201c and \u201d = The weird formatted quote marks
        var group = 1; // Item in parentheses
        try {
          description = content.match(regex)[group];
        } catch {
          // Description not found.
        }
        break;
      }
      case (PaymentSource.BRCH3Website): {
        var regex = "Product Quantity Price\\r?\\n((.+\\r?\\n)+)Subtotal";
        var group = 1;
        try {
          description = content.match(regex)[group];
        } catch {
          // Description not found.
        }
      }
      case (PaymentSource.PayPal): {        
        var regex = "\\[image: quote](\\r?\\n)+(.*)(\\r?\\n)+\\[image: quote]";
        var group = 2;
        try {
          description = content.match(regex)[group];
        } catch {
          // Description not found.
        }
      }
    }
    return description;
  }
  
  getPaymentAmount() {
    let content = this.emailContent;

    switch (this.paymentSource) {
      case (PaymentSource.Zelle): {
        var regex = "Amount *.(.*)";
        var group = 1; // Item in parentheses, e.g., "$250.00"
        break;
      }
      case (PaymentSource.GPay): {
        var regex = "(\\$[0-9]+\\.*[0-9]*)"; // Looks for dollar sign amount
        var group = 1; // Item in parentheses, e.g., "$250.00"
        break;
      }
      case (PaymentSource.BRCH3Website): {
        var regex = "Total: (\\$[0-9]+\\.[0-9]{2})";
        var group = 1; // Item in parentheses, e.g., "$25.00"
        break;
      }
      case (PaymentSource.PayPal): {   
        var regex = "sent you (.*) USD";
        var group = 1; // Item in parentheses, e.g., "$25.00"
        break;
      }
    }
    return content.match(regex)[group];
  }
  
  getPaymentDate() {
    let content = this.emailContent;
    let variableDateRegex = "(\\b\\d{1,2}\\D{0,3})?\\b(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|(Nov|Dec)(?:ember)?)\\D?(\\d{1,2}\\D?)?\\D?((19[7-9]\\d|20\\d{2})|\\d{2})";

    switch (this.paymentSource) {
      case (PaymentSource.Zelle): {
        var regex = "Date \\*{1}\\s*(.*)";
        var group = 1; // Item in parentheses
        break;
      }
      case (PaymentSource.GPay): {
        var regex = variableDateRegex;
        var group = 0; // Full date match, e.g., "March 14, 2023"
        break;
      }
      case (PaymentSource.BRCH3Website): {
        var regex = variableDateRegex;
        var group = 0; // Full date match, e.g., "March 14, 2023"
        break;
      }
      case (PaymentSource.PayPal): {  
        var regex = "\\*Transaction\\r?\\ndate\\*\\r?\\n(.*)\\r?\\n";
        var group = 1; // Item in parentheses
        break;
      }
    }
    return formatDate(content.match(regex)[group]);
  }

  getCamperName() {
    let content = this.emailContent;

    switch (this.paymentSource) {
      case (PaymentSource.Zelle): {
        var regex = "(.*) sent you money";
        var group = 1; // Item in parentheses
        break;
      }
      case (PaymentSource.GPay): {
        var regex = "(.*) sent you money";
        var group = 1; // Item in parentheses
        break;
      }
      case (PaymentSource.BRCH3Website): {
        var regex = "order from (.*):";
        var group = 1; // Item in parentheses
        break;
      }
      case (PaymentSource.PayPal): { 
        var regex = "(.*) sent you \\$";
        var group = 1; // Item in parentheses
        break;
      }
    }
    return content.match(regex)[group];
  }

  getPaymentDue(earliestPaymentDate) {   
      let totalDue = calculateTotalDues(earliestPaymentDate, this.camperNames.hashName);
    
      // Return to current Payments sheet      
      setActiveSpreadsheet(Definitions.paymentsTabName);

      return formatCurrency(totalDue);
  }
  
  //---------------------------------------------------------------------------------------
  // Look at all existing rows matching camper name for total amount paid and required dues
  //---------------------------------------------------------------------------------------
  makeMultipleRowCalculations() {
    var tab = Definitions.paymentsTabName;
    var sheet = setActiveSpreadsheet(tab);    
    var hashNameHeaderCol = Columns.hashName - 1;
    var paymentAmountHeaderCol = Columns.paymentAmount - 1;
    var paymentDateHeaderCol = Columns.paymentDate - 1;

    let requiredDues = 0;
    let totalPaid = parseCurrency(this.paymentAmount);
    let earliestCampPaymentDate = this.paymentDate;

    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    var firstRow = Rows.firstPaymentsDataRow;
      
    let hasherName = this.camperNames.hashName;
    for (let i = firstRow; i < values.length; i++) {      
      let rowHasherName = values[i][hashNameHeaderCol];

      if (rowHasherName == hasherName && hasherName != "?") {
        let paymentDateCell = formatDate(values[i][paymentDateHeaderCol].toString());
        let earliestPayment = formatDate(findEarlierDate(earliestCampPaymentDate, paymentDateCell));

        let paymentAmountCell = values[i][paymentAmountHeaderCol].toString();          
        let rowPaid = parseCurrency(paymentAmountCell);
        let newSum = rowPaid + totalPaid;
        totalPaid = newSum;

        requiredDues = calculateRequiredDues(earliestPayment);
        if (totalPaid >= parseCurrency(requiredDues.toString())) { 
          earliestCampPaymentDate = earliestPayment;
        }
      }
    } 
    requiredDues = calculateRequiredDues(earliestCampPaymentDate);
    
    return { earliestPaymentDate: earliestCampPaymentDate, totalPaid: formatCurrency(totalPaid), requiredDues: requiredDues };
  }

  //--------------------------------------------------------------------
  // For wonky hash name definitions from camper registration form. 
  // Tries to match partial names by splitting up registered name into parts.
  //--------------------------------------------------------------------
  tryComplexCamperName(registeredHashName, paymentDescription) {    
    var foundHashName = "";
    var registeredHashNameParts = registeredHashName.split(" ");
    var descriptionParts = paymentDescription.split(" ");

    for(var i = 0; i < registeredHashNameParts.length; i++) {
      let word = registeredHashNameParts[i];
      let parenMatch = "(" + word + ")";
      if (descriptionParts.includes(word) 
        || descriptionParts.includes(parenMatch)) {
        foundHashName += word + " ";
      }
    }
    if (foundHashName.trim().toUpperCase() === "JUST") {
      return "";
    }
    return foundHashName.trim();
  }
}
