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
    this.paymentDue = this.getPaymentDue(this.paymentDate);
    this.paymentAmountTotal = this.getAllPaymentsMade();

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
          var foundHashName = hasHashNameToFind ? paymentDescription.toUpperCase().includes(hashName.toUpperCase())
                                                : false;   
          var foundComplexHashName = hasHashNameToFind ? this.tryComplexCamperName(hashName, paymentDescription)
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
          var hasFirstAndLastNameInMoneySender = moneySender.includes(" ") && moneySender.length > 1;
          var hasNameInMoneySender = moneySender.length > 1;
          var foundMultipleCandidates = false;
          var candidate = null;

          for (var i = 1; i < values.length; i++) 
          {
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
    if (this.emailContent.toUpperCase().indexOf("ZELLE") > -1) {
      return PaymentSource.Zelle;
    }
    if (this.emailContent.toUpperCase().indexOf("A WANKER GAVE US MONEY") > -1) {
      return PaymentSource.BRCH3Website;
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
    }
    return content.match(regex)[group];
  }

  getPaymentDue(paymentDate) {   
      let totalDue = 0; 
      var tab = Definitions.habOrdersTabName;
      var habOrdersSheet = setActiveSpreadsheet(tab);
      
      var hashNameHeaderCol = getByRangeName(tab, 'HabOrders.HashName').column - 1;
      var totalDueHeaderCol = getByRangeName(tab, 'HabOrders.TotalDue').column - 1;

      var dataRange = habOrdersSheet.getDataRange();
      var values = dataRange.getValues();
      var firstRow = Rows.paymentDueDataRow + 1;
      
      for (let i = firstRow; i < values.length; i++) 
      {
        let rowHasherName = values[i][hashNameHeaderCol];

        if (rowHasherName == this.camperNames.hashName) {
          let totalDueCell = values[i][totalDueHeaderCol];
          let multipleAmountRegex = new RegExp(/(.*)(Amount: (.*) USD)/g);
          let descriptionGroup = 1;
          let amountGroup = 3;
          var result = null;

          while((result = multipleAmountRegex.exec(totalDueCell)) !== null) {          
            let description = result[descriptionGroup];
            let amount = result[amountGroup];

            if (!description.toUpperCase().includes("CAMP DUES")) {
              totalDue += parseCurrency(amount);
            }
          }

          let requiredDues = calculateRequiredDues(paymentDate);
          if (Number.isInteger(requiredDues)) {
            totalDue += requiredDues;
          }
        }
      }
    
      // Return to current Payments sheet      
      setActiveSpreadsheet(Definitions.paymentsTabName);

      return formatCurrency(totalDue);
  }
  
  //--------------------------------------------------------------------
  // Add up all existing rows matching camper name for total amount paid
  //--------------------------------------------------------------------
  getAllPaymentsMade() {
    var tab = Definitions.paymentsTabName;
    var sheet = setActiveSpreadsheet(tab);    
    var hashNameHeaderCol = getByRangeName(tab, 'ScrapedEmailData.HashName').column - 1;
    var paymentAmountHeaderCol = getByRangeName(tab, 'ScrapedEmailData.PaymentAmount').column - 1;

    let totalPaid = parseCurrency(this.paymentAmount);
    let hasherNames = [ this.camperNames.hashName ];

    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    var firstRow = Rows.firstPaymentsDataRow;
      
    for (let j = 0; j < hasherNames.length; j++) {
      let hasherName = hasherNames[j];
      for (let i = firstRow; i < values.length; i++) 
      {      
        let rowHasherName = values[i][hashNameHeaderCol];

        if (rowHasherName == hasherName) {
          let paymentAmountCell = values[i][paymentAmountHeaderCol].toString();          
          let rowPaid = parseCurrency(paymentAmountCell);
          let newSum = rowPaid + totalPaid;
          totalPaid = newSum;
        }
      }      
    }
    return formatCurrency(totalPaid);
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
      if (descriptionParts.includes(word)) {
        foundHashName += word + " ";
      }
    }
    return foundHashName.trim();
  }
}
