function import_missing_accounts() {   

  // Looks for domains that don't have an account, presumably because
  // the domain is from an account external to the region being processed.
  // Also tracks emails to find leads should the missing account not exist.
  
  logStamp("Missing Account Import");

  
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
  let missingAccountInfo = load_tab_(MISSING_DOMAINS, 2, 1);
  for (var i=0; i< missingAccountInfo.length; i++) {
    missingAccounts[missingAccountInfo[i][MISSING_EMAIL_DOMAIN]] = true;
  }
  
  
  // 
  // Load current catalog of potential leads (based on missing accounts)
  //
  let potentialLeads = load_tab_(MISSING_EMAILS, 2, 1);
  for (var i=0; i< potentialLeads.length; i++) {
    potentialLeadEmails[potentialLeads[i][0]] = true;
  }
  
  // 
  // Process Calendar invites
  //
  
  // The raw calendar invites are in the Calendar tab.  
  let inviteInfo = load_tab_(CALENDAR, 2, CALENDAR_COLUMNS);
  
  Logger.log(CALENDAR + " import. Size is " + inviteInfo.length);
  
  if (inviteInfo.length == 0) return; // Empty (or error). Only header
  
  // Set missing domain output cursor
  sheet = SpreadsheetApp.getActive().getSheetByName(MISSING_DOMAINS);
  let elr = sheet.getLastRow();
  let outputRange = sheet.getRange(elr+1,1);
  let domainOutputCursor = {range : outputRange, rowOffset : 0};
  
    // Set lead output cursor
  sheet = SpreadsheetApp.getActive().getSheetByName(MISSING_EMAILS);
  elr = sheet.getLastRow();
  let leadOutputRange = sheet.getRange(elr+1,1);
  let leadOutputCursor = {range : leadOutputRange, rowOffset : 0};
  
  // 
  // Look for email domains that don't have an In Region Customer or Partner account
  //
  
  let detectedAccounts = {}; // For the new stuff
  let loggedLeads = {}; // For the new stuff
  for (var j = 0 ; j < inviteInfo.length; j++) {
    
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
    // attendeeInfo.stats.others - Number of unidentified attendees
    
    
    // If no customers or partners, track the missing domains so we can lookup up accounts later from Salesforce
    // (We can't pull in every account in Salesforce because there are too many. So, we look up stuff that isn't
    // "In Region".)
    if (attendeeInfo.stats.customers == 0 && attendeeInfo.stats.partners == 0) {
      for (d in attendeeInfo.others) {
        
        for (let j=0; j<attendees.length; j++) {
          
          let attendeeEmail = attendees[j];
          
          // By definition, only emails for domains not yet associated with an account are in the attendee list.
          // However, there may be some hashicorp folks. Filter them out.
          // Record all of these non-hashi emails as potential leads (if an accociated account isn't in Salesforce yet)
          
          if (potentialLeadEmails[attendeeEmail] || loggedLeads[attendeeEmail]) {
            continue;
          }
          if (attendeeEmail.indexOf("hashicorp") != -1) continue;
          
          loggedLeads[attendeeEmail] = true;
          leadOutputCursor.range.offset(leadOutputCursor.rowOffset, 0).setValue(attendeeEmail);     
          leadOutputCursor.rowOffset++;   
          
        }     
        
        if (missingAccounts[d] || detectedAccounts[d]) {
          continue; // old news
        }
        
        detectedAccounts[d] = true;
        domainOutputCursor.range.offset(domainOutputCursor.rowOffset, MISSING_EMAIL_DOMAIN).setValue(d);     
        domainOutputCursor.rowOffset++;     
        
        
      }
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

  logStamp("Staging Upload");
  
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
  let uc = 0;
  
  
  for (j = 0 ; j < elr - 1; j++) {
  
    if (eventInfo[j][EVENT_PROCESS] == PROCESS_SKIP)  {
      let date = Utilities.formatDate(new Date(eventInfo[j][EVENT_START]), "GMT-5", "MMM dd, yyyy");
      logOneCol("Skipped " + eventInfo[j][EVENT_SUBJECT] + " / " + date);
      continue;
    }
    
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
    uploadRange.offset(rowOffset, EVENT_LEAD).setValue(eventInfo[j][EVENT_LEAD]); 
    uc++;
    rowOffset++;
  }
  logOneCol("Uploaded " + uc + " events.");
  
}

function process_account_emails_(accountInfo, numberOfRows, accountType, accountName, accountId, emailFieldNumber, altEmailFieldNumber, emailToAccountMap) {
  // accountInfo is an array of account records. numberOfRows is the number of accounts. All the other parameters, with the exception 
  // of the last, are the "field numbers" (not sure why I didn't name all of them as <field>FieldNumber). The last parameter, emailToAccountMap, 
  // is an object to record what email domain belongs to what account.
  
  // This normally process email domains (on stuff after the @ sign). However, one of the callers will pass in a full email (<x>@<domain>). So, make
  // sure the logic is generic enough to handle that.
  
  let accountLog = {};
  
  // Build the email domain to customer mapping  
  for (var j = 0 ; j < numberOfRows; j++) {
  
    let emailDomainString = accountInfo[j][emailFieldNumber];
    let altDomainString = 0;
    if (altEmailFieldNumber) {
      altDomainString = accountInfo[j][altEmailFieldNumber];  
    }  
  
    if (!emailDomainString && !altDomainString) {
      Logger.log("WARNING :" + accountInfo[j][accountName] + " has no email domain!");
      continue;
    }
    let emailDomains = [];
    if (emailDomainString && typeof emailDomainString == "string") {
      emailDomains = emailDomainString.split(','); // Works if there is only one (no comma)
    }
    if (altDomainString && typeof altDomainString == "string") {    
      let moreDomains = altDomainString.split(','); // 99% of the time, domains are comma separated, but not always
      if (moreDomains.length > 0) {
        emailDomains = emailDomains.concat(moreDomains);
      }
    }
    
  
    let uniqueDomains = {};
    
    for (k = 0; k < emailDomains.length; k++) {
      // LATAM domains violate rules
      //let emailRegex = /[-\w]+\.[a-zA-Z]{2,3}$/;
      //let domain = emailRegex.exec(emailDomains[k].trim());
      let superDomain = emailDomains[k].trim(); // probably just one domain, but you never know
      
      // Because sometimes Reps use spaces instead of commas to separate domains in
      // the salesforce field. Deal with it now.
      let potentiallyMoreDomains = superDomain.split(" ");
      for (let i = 0; i < potentiallyMoreDomains.length; i++) {
      
        let domain = potentiallyMoreDomains[i].trim();
        
        // FIXME other prefixes I don't know about? 
        if (domain.indexOf("www.") == 0) {
          domain = domain.substring(4);
        }
        
        if (uniqueDomains[domain]) {
          continue; // Rep put this domain in twice!
        }
        uniqueDomains[domain] = true;
        
        if (accountLog[domain]) {
          //Logger.log("WARNING: Found account with a duplicate email domain - " + accountInfo[j][accountName] + ":" + domain);
          
          logFourCol("Multiple accounts for email domain:", domain, "Ignored: " + accountInfo[j][accountName], "Selected: " + accountLog[domain]);         
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
  
  sheet = SpreadsheetApp.getActive().getSheetByName(MISSING_CUSTOMERS);
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
  
  process_account_emails_(partnerInfo, plr - 1, PARTNER_TYPE, PARTNER_NAME, PARTNER_ID, PARTNER_EMAIL_DOMAIN, 0, emailToPartnerMap); // TODO - add PARTNER_ALT_EMAIL_DOMAINS into partner report
  
  return true;
}

function load_lead_info_() {
 
  let leadInfo = load_tab_(MISSING_LEADS, 2, 1);
  if (leadInfo.length > 0) {
    process_account_emails_(leadInfo, leadInfo.length, LEAD_TYPE, LEAD_NAME, LEAD_ID, LEAD_EMAIL, 0, emailToLeadMap);
  }
}


function printIt_() {
  let keys = PropertiesService.getScriptProperties().getKeys();
  for (var i = 0; i < keys.length; i++) {
    Logger.log(keys[i] + ":" + PropertiesService.getScriptProperties().getProperty(keys[i]));
  }
}

function findLead_(attendees) {
  
  for (let i=0; i<attendees.length; i++) {  
    if (emailToLeadMap[attendees[i]]) {
      return emailToLeadMap[attendees[i]];
    }
  }
  return null;
}

function load_tab_(sheetName, fromRow, minColumnCount) { 
  
  try {
    let sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    let rangeData = sheet.getDataRange();
    let lastColumn = rangeData.getLastColumn();
    let lastRow = rangeData.getLastRow();
    
    if (lastColumn < minColumnCount) {
      Logger.log("ERROR: " + sheetName + " does not have " + minColumnCount + " fields.");
      return [];
    }
    
    if (lastRow >= fromRow) {
      let rowCount = lastRow-fromRow+1
      let scanRange = sheet.getRange(fromRow,1, rowCount, lastColumn);
      let inputArray = scanRange.getValues();
      return inputArray;
    }
  }
  catch (e) {
    Logger.log("load_tab_ exception: " + e);
  }
  
  return [];
}

function generateLongId_(id) {

  // Transform a 15 character case sensitive Salesforce ID into the 
  // "long version", an 18 character case insensitive id.
  // This is Salesforce's algorithm.
  
  let retVal = id;
  
  if(id.length == 15){
    
    let addon="";
    for(let block=0;block<3; block++)
    {
      let loop=0;
      for(let position=0;position<5;position++){
        let current=id.charAt(block*5+position);
        if(current>="A" && current<="Z")
          loop+=1<<position;
      }
      addon+="ABCDEFGHIJKLMNOPQRSTUVWXYZ012345".charAt(loop);
    }
    retVal=(id+addon);
  }
  return retVal;
}

function logStamp(title) {
  
  let logSheet = SpreadsheetApp.getActive().getSheetByName(LOG_TAB);
  var logLastRow = logSheet.getLastRow();
  AM_LOG = logSheet.getRange(logLastRow+2,1); // Leave an empty row divider
  AM_LOG_ROW = 0;
  
  AM_LOG.offset(AM_LOG_ROW, 0).setValue(title + " " + new Date().toLocaleTimeString());
  AM_LOG.offset(AM_LOG_ROW+1, 0).setValue("---------------------------------------------------------------");
  AM_LOG_ROW+=2; 
}

function logOneCol(message) {
  AM_LOG.offset(AM_LOG_ROW, 0).setValue(message);
  AM_LOG_ROW++;   
}

function logTwoCol(one, two) {
  AM_LOG.offset(AM_LOG_ROW, 0).setValue(one);
  AM_LOG.offset(AM_LOG_ROW, 1).setValue(two);
  AM_LOG_ROW++;   
}

function logThreeCol(one, two, three) {
  AM_LOG.offset(AM_LOG_ROW, 0).setValue(one);
  AM_LOG.offset(AM_LOG_ROW, 1).setValue(two);
  AM_LOG.offset(AM_LOG_ROW, 2).setValue(three);
  AM_LOG_ROW++;   
}

function logFourCol(one, two, three, four) {
  AM_LOG.offset(AM_LOG_ROW, 0).setValue(one);
  AM_LOG.offset(AM_LOG_ROW, 1).setValue(two);
  AM_LOG.offset(AM_LOG_ROW, 2).setValue(three);
  AM_LOG.offset(AM_LOG_ROW, 3).setValue(four);
  AM_LOG_ROW++;    
}