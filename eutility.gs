function import_external_accounts() {    
  
  //
  // Load Account Info
  //
  
  if (!load_customer_info_()) {
    return;
  }
  if (!load_partner_info_()) {
    return;
  }
  
  // 
  // Load current catalog of (formerly) missing accounts
  //
  let sheet = SpreadsheetApp.getActive().getSheetByName(MISSING_CUSTOMERS);
  let rangeData = sheet.getDataRange();
  let lastColumn = rangeData.getLastColumn();
  let lastRow = rangeData.getLastRow();
  
  if (lastRow > 1) {
    let scanRange = sheet.getRange(2,1, lastRow-1, lastColumn);
    
    // Suck all the calendar data up into memory.
    let missingAccountInfo = scanRange.getValues();
    for (var i=0; i< lastRow -1; i++) {
      missingAccounts[missingAccountInfo[i][MISSING_EMAIL_DOMAIN]] = true;
    }
  }
  
  // 
  // Process Calendar invites
  //
  
  // The raw calendar invites are in the Calendar tab.
  sheet = SpreadsheetApp.getActive().getSheetByName('Calendar')
  rangeData = sheet.getDataRange();
  lastColumn = rangeData.getLastColumn();
  lastRow = rangeData.getLastRow();
  
  if (lastRow == 1) return; // Empty. Only header
  
  if (lastColumn < CALENDAR_COLUMNS){
    Logger.log("ERROR: Imported Calendar was only " + lastColumn + " fields wide. Not enought! Something bad happened.");
    return;
  }
  scanRange = sheet.getRange(2,1, lastRow-1, lastColumn);
  
  // Suck all the calendar data up into memory.
  let inviteInfo = scanRange.getValues();
  
  // Set missing domain output cursor
  sheet = SpreadsheetApp.getActive().getSheetByName(MISSING_CUSTOMERS);
  let elr = sheet.getLastRow();
  let outputRange = sheet.getRange(elr+1,1);
  let outputCursor = {range : outputRange, rowOffset : 0};
  
  // 
  // Convert calendar invites (Calendart tab) to SE events (Events tab)
  //
  
  let detectedAccounts = {};
  for (var j = 0 ; j < lastRow - 1; j++) {
    
    if (!inviteInfo[j][ASSIGNED_TO] || !inviteInfo[j][ATTENDEE_STR] || !inviteInfo[j][START]) {
      continue;
    }
    
    var attendees = inviteInfo[j][ATTENDEE_STR].split(","); // convert comma separated list of emails to an array
    if (attendees.length == 0) {
      continue; 
    }
    
    //
    // Determine who was at the meeting
    //
    
    var attendeeInfo = lookForAccounts_(attendees, emailToCustomerMap, emailToPartnerMap);
    // attendeeInfo.customers - Map of prospect/customer Salesforce 18-Digit account id to number of attendees
    // attendeeInfo.partners - Map of partner Salesforce 18-Digit account id to number of attendees
    // attendeeInfo.others - Map of unknown email domains to number of attendees
    // attendeeInfo.stats.customers - Number of customers in attendence
    // attendeeInfo.stats.partners - Number of partners in attendence
    // attendeeInfo.stats.hashi - Number of hashicorp attendees
    // attendeeInfo.stats.other - Number of unidentified attendees
    
    for (d in attendeeInfo.others) {

      if (missingAccounts[d] || detectedAccounts[d]) {
        continue; // old news
      }
      
      detectedAccounts[d] = true;
      outputCursor.range.offset(outputCursor.rowOffset, MISSING_EMAIL_DOMAIN).setValue(d);     
      outputCursor.rowOffset++;     
    }
  }
  // Salesforce OAUTH2 doesn't work, so we have to use Zapier
  // var html = HtmlService.createTemplateFromFile('index').evaluate().setWidth(500);
  // SpreadsheetApp.getUi().showSidebar(html);
}



//function onOpen() {
//  Browser.msgBox('App Instructions - Please Read This Message', '1) Click Tools then Script Editor\\n2) Read/update the code with your desired values.\\n3) Then when ready click Run export_gcal_to_gsheet from the script editor.', Browser.Buttons.OK);
//}


function getProducts_(productScan) {
  var products = []
  if (productScan.hasTerraform) {
    products.push("Terraform");
  }
  if (productScan.hasVault) {
    products.push("Vault");
  }
  if (productScan.hasConsul) {
    products.push("Consul");
  }
  if (productScan.hasNomad) {
    products.push("Nomad");
  }
  if (products.length == 0) {
    return "None";
  }
  else {
    return products.toString();
  }
}

function getOneProduct_(productScan) {
  if (productScan.hasTerraform) {
    return "Terraform";
  }
  if (productScan.hasVault) {
    return "Vault";
  }
  if (productScan.hasConsul) {
    return "Consul";
  }
  if (productScan.hasNomad) {
    return "Nomad";
  }
  return "N/A";
}

function makeProductKey_(productScan, primary) {
  
  let returnValue = "-";
  
  if (primary && primary != "Terraform" && primary != "Vault" && primary != "Consul" && primary != "Nomad") {
    let x = primary.toLowerCase();
    if (x.indexOf("terraform") != -1) primary = "Terraform";
    else if (x.indexOf("tfe") != -1) primary = "Terraform";
    else if (x.indexOf("tfc") != -1) primary = "Terraform";
    else if (x.indexOf("v") != -1) primary = "Vault";
    else if (x.indexOf("c") != -1) primary = "Consul";
    else if (x.indexOf("n") != -1) primary = "Nomad";
    else Logger.log("WARNING: " + primary + "is not a known primary product!");
  }
  
  if (productScan.hasTerraform || (primary && (primary == "Terraform"))) returnValue += TERRAFORM;
  if (productScan.hasVault || (primary && primary == "Vault")) returnValue += VAULT;
  if (productScan.hasConsul || (primary && primary == "Consul")) returnValue += CONSUL;
  if (productScan.hasNomad || (primary && primary == "Nomad")) returnValue += NOMAD;
  
  return returnValue;
  
}

function makeSingleProductKeys_(productScan) {

  // Function is used to create "individual product" keys. If a calendar
  // invite specifies multiple products, but there are no opportunities for 
  // that combination of products, we want to look for opportunities that may
  // have one of the products. So, we need to create an array of possible keys
  
  let retval = []
  let i = 0;
  
  if (productScan.hasTerraform) retval[i++] = "-" + TERRAFORM;
  if (productScan.hasVault) retval[i++] = "-" + VAULT;
  if (productScan.hasConsul) retval[i++] = "-" + CONSUL;
  if (productScan.hasNomad) retval[i++] = "-" + NOMAD;
  
  return retval;
  
}

/*
function testregex() {
  
  let text = "Prep:1DhYes";
  let regex = /[Pp][Rr][Ee][Pp] *: *[0-9]+[ ]?[MmHhDd]?/; 
  let prepArray = text.match(regex);
  if (prepArray && prepArray[0]) {
    let kv = prepArray[0].split(':');
    let prep = parseInt(kv[1]); // minutes
    let units = kv[1].match(/[MmHhDd]$/);
    if (units) {
      if (units == "h" || units == "H") {
        prep = prep * 60;
      }
      else if (units == "d" || units == "D") {
        prep = prep * 60 * 8; // 8 hour workday
      }
      Logger.log("Prep: " + prep);
    }
  }
}
*/




function clearTab_(tab_name, header) {
  
  if (!tab_name) return;
  
  let sheet = SpreadsheetApp.getActive().getSheetByName(tab_name);
  sheet.clearContents();  
  
  if (header) {
    let range = sheet.getRange(1,1,1,header[0].length);
    range.setValues(header);
  }
}


function stage_events_to_upload_tab() {
  
  // Do NOT clear out old records! Zappier needs to be off before doing that!
  
  let sheet = SpreadsheetApp.getActive().getSheetByName(EVENTS);
  let rangeData = sheet.getDataRange();
  let elc = rangeData.getLastColumn();
  let elr = rangeData.getLastRow();
  if (elr == 1) return; // Events is empty! Only has the header.
  let scanRange = sheet.getRange(2,1, elr-1, elc); // 2 is to skip header. Assumes a header!!!!
  let eventInfo = scanRange.getValues();
  
  if (elc < EVENT_COLUMNS) {
    Logger.log("ERROR: Imported Customers was only " + elc + " fields wide. Not enough! Something is wrong.");
    return;
  }
  
  // Set event output cursor
  let sheet2 = SpreadsheetApp.getActive().getSheetByName(UPLOAD_STAGE);
  let ulr = sheet2.getLastRow();
  let uploadRange = sheet2.getRange(ulr+1,1);
  let rowOffset = 0;
  
  
  for (j = 0 ; j < elr - 1; j++) {
    uploadRange.offset(rowOffset, EVENT_ASSIGNED_TO).setValue(eventInfo[j][EVENT_ASSIGNED_TO]);
    uploadRange.offset(rowOffset, EVENT_OP_STAGE).setValue(eventInfo[j][EVENT_OP_STAGE]);
    uploadRange.offset(rowOffset, EVENT_MEETING_TYPE).setValue(eventInfo[j][EVENT_MEETING_TYPE]);
    uploadRange.offset(rowOffset, EVENT_RELATED_TO).setValue(eventInfo[j][EVENT_RELATED_TO]);
    uploadRange.offset(rowOffset, EVENT_SUBJECT).setValue(eventInfo[j][EVENT_SUBJECT]);
    uploadRange.offset(rowOffset, EVENT_START).setValue(eventInfo[j][EVENT_START]);
    uploadRange.offset(rowOffset, EVENT_END).setValue(eventInfo[j][EVENT_END]);
    uploadRange.offset(rowOffset, EVENT_REP_ATTENDED).setValue(eventInfo[j][EVENT_REP_ATTENDED]);
    uploadRange.offset(rowOffset, EVENT_PRODUCT).setValue(eventInfo[j][EVENT_PRODUCT]);
    uploadRange.offset(rowOffset, EVENT_DESC).setValue(eventInfo[j][EVENT_DESC]);
    uploadRange.offset(rowOffset, EVENT_LOGISTICS).setValue(eventInfo[j][EVENT_LOGISTICS]);
    uploadRange.offset(rowOffset, EVENT_PREP_TIME).setValue(eventInfo[j][EVENT_PREP_TIME]);  
    uploadRange.offset(rowOffset, EVENT_QUALITY).setValue(eventInfo[j][EVENT_QUALITY]);  
    
    
    rowOffset++;
  }
}

function process_account_emails_(accountInfo, numberOfRows, accountType, accountName, accountId, emailFieldNumber, altEmailFieldNumber, emailToAccountMap) {
  
  let accountLog = {};
  
  // Build the email domain to customer mapping  
  for (var j = 0 ; j < numberOfRows; j++) {
  
    let emailDomainString = accountInfo[j][emailFieldNumber];
    if (!emailDomainString) {
      Logger.log("WARNING :" + accountInfo[j][accountName] + " has no email domain!");
      continue;
    }
    if (typeof emailDomainString != "string") {
      Logger.log("WARNING :" + accountInfo[j][accountName] + " has a non-string email domain: " + emailDomainString);
      continue;
    }
    
    let emailDomains = emailDomainString.split(','); // Works if there is only one (no comma)
    
    let altDomainString = accountInfo[j][altEmailFieldNumber];  
    if (altDomainString) {
       if (typeof altDomainString == "string") {     
        let moreDomains = altDomainString.split(',');
        if (moreDomains.length > 0) {
          emailDomains = emailDomains.concat(moreDomains);
        }
      }
    }
    
    for (var k = 0; k < emailDomains.length; k++) {
      // LATAM domains violate rules
      //let emailRegex = /[-\w]+\.[a-zA-Z]{2,3}$/;
      //let domain = emailRegex.exec(emailDomains[k].trim());
      let domain = emailDomains[k].trim();
      // FIXME other prefixes I don't know about? 
      if (domain.indexOf("www.") == 0) {
        domain = domain.substring(4);
      }
      //DAK
      if (accountLog[domain]) {
        //Logger.log("WARNING: Found account with a duplicate email domain - " + accountInfo[j][accountName] + ":" + domain);
        
        AM_LOG.offset(AM_LOG_ROW, 0).setValue("Multiple accounts for email domain:");
        AM_LOG.offset(AM_LOG_ROW, 1).setValue(domain);
        AM_LOG.offset(AM_LOG_ROW, 2).setValue("Rejected: " + accountInfo[j][accountName]);
        AM_LOG.offset(AM_LOG_ROW, 3).setValue("Selected: " + accountLog[domain]);
        AM_LOG_ROW++;    
        
        continue; // Take the first
      }
      accountLog[domain] = accountInfo[j][accountName];
      let id = accountInfo[j][accountId].trim();
      emailToAccountMap[domain] = id;
      accountType[id] = accountType;
      //Logger.log("DEBUG: " + domain + " -> " + emailToAccountMap[domain]);
    }
  }
}

function load_customer_info_() {
  
  //
  // Load up our region's customers
  //

  let sheet = SpreadsheetApp.getActive().getSheetByName(IN_REGION_CUSTOMERS);
  
  let rangeData = sheet.getDataRange();
  let alc = rangeData.getLastColumn();
  let alr = rangeData.getLastRow();
  if (alr < 2) {
    // No partners (that's odd)
    Logger.log("ERROR: No Customers found! Perhaps you should refresh the In Region Customer tab.");
    return false;
  }
  let scanRange = sheet.getRange(2,1, alr-1, alc);
  let customerInfo = scanRange.getValues();
  
  if (alc < CUSTOMER_COLUMNS) {
    Logger.log("ERROR: Imported Customers was only " + alc + " fields wide. Not enough! Something is wrong.");
    return false;
  }
  
  process_account_emails_(customerInfo, alr - 1, INTERNAL_CUSTOMER_TYPE, CUSTOMER_NAME, CUSTOMER_ID, CUSTOMER_EMAIL_DOMAIN, CUSTOMER_ALT_EMAIL_DOMAINS, emailToCustomerMap);
  
  //
  // Load up customers outside of our region
  //
  
  sheet = SpreadsheetApp.getActive().getSheetByName(EXTERNAL_CUSTOMERS);
  rangeData = sheet.getDataRange();
  alc = rangeData.getLastColumn();
  alr = rangeData.getLastRow();
  if (alr < 2) {
    // No external customers logged
    return true;
  }
  scanRange = sheet.getRange(2,1, alr-1, alc);
  customerInfo = scanRange.getValues();
  
  if (alc < CUSTOMER_COLUMNS) {
    Logger.log("ERROR: Imported Customers was only " + alc + " fields wide. Not enough! Something is wrong.");
    return true; // run with the internal customers
  }
  
  process_account_emails_(customerInfo, alr - 1, EXTERNAL_CUSTOMER_TYPE, CUSTOMER_NAME, CUSTOMER_ID, CUSTOMER_EMAIL_DOMAIN, CUSTOMER_ALT_EMAIL_DOMAINS, emailToCustomerMap);
   
  return true;
}

function load_partner_info_() {
  
  let sheet = SpreadsheetApp.getActive().getSheetByName(PARTNERS);
  let rangeData = sheet.getDataRange();
  let plc = rangeData.getLastColumn();
  let plr = rangeData.getLastRow();
  if (plr < 2) {
    // No partners (that's odd)
    Logger.log("ERROR: No Partners found! Perhaps you should refresh the partner tab.");
    return false;
  }
  let scanRange = sheet.getRange(2,1, plr-1, plc);
  let partnerInfo = scanRange.getValues();
  
  if (plc < PARTNER_COLUMNS) {
    Logger.log("ERROR: Imported Partners was only " + plc + " fields wide. Not enough! Something is awry.");
    return false;
  }
  
  process_account_emails_(partnerInfo, plr - 1, PARTNER_TYPE, PARTNER_NAME, PARTNER_ID, PARTNER_EMAIL_DOMAIN, PARTNER_ALT_EMAIL_DOMAINS, emailToPartnerMap);
  
  return true;
}


function printIt_() {
  let keys = PropertiesService.getScriptProperties().getKeys();
  for (var i = 0; i < keys.length; i++) {
    Logger.log(keys[i] + ":" + PropertiesService.getScriptProperties().getProperty(keys[i]));
  }
}
