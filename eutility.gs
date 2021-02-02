function createDataLoadFilters() {

  // Looks for domains that don't have an in-region account, presumably because
  // the domain is from an account external to the region being processed, or
  // the account doesn't exist. These "out-region" or missing domains are tracked in ...
  //     detectedDomains
  // Also tracks emails to find leads should the out-of-region account not exist.

  logStamp("Create Data Load Filters");

  // Identifying out-of-region customers and relevant leads to filter the Customer and Leads tabs during event processing

  //
  // Load Account Info
  //

  let partnerMap = {}; // domain to account
  let customerMap = {};
  let typeMap = {}; // account to account type
  // These maps will be UNFILTERED (every record will be in them. Not a problem for in-region customers, but the partner list is pretty big.
  // We only need these full lists temporarily in this function, so we are not using the global emailToCustomerMapG and emailToPartnerMapG variables.
  // Hopefully these temporary maps will be garbage collected after this function finishes. 

  // Load in-region customers only, the first parm set to true (we know that list won't be massive)
  // Also updates accountType object
  if (!load_customer_info_(true, customerMap, typeMap, 1, true)) {
    return;
  }

  // Loads ALL partners
  // Saves info off in 
  // Also updates accountType object
  if (!load_partner_info_(true, partnerMap, typeMap, 1, true)) {
    return;
  }

  // At this point, all the "known" accounts are stored in the accountType object. 

  clearTab_(MISSING_DOMAINS, [['Email']]);
  clearTab_(MISSING_CUSTOMERS, [['Account Id']]);
  clearTab_(POTENTIAL_LEADS, [['Email']]);
  clearTab_(IN_PLAY_CUSTOMERS, [['Account Id']]);
  clearTab_(IN_PLAY_PARTNERS, [['Account Id']]);

  // 
  // Process Calendar invites
  //

  // The raw calendar invites are in the Calendar tab.  
  let inviteInfo = load_tab_(CALENDAR, 2, CALENDAR_COLUMNS);

  //Logger.log(CALENDAR + " import. Size is " + inviteInfo.length);

  if (inviteInfo.length == 0) return; // Empty (or error). Only header

  // Set missing domain output cursor
  let sheet = SpreadsheetApp.getActive().getSheetByName(MISSING_DOMAINS);
  let lr = sheet.getLastRow();
  let missingOutputRange = sheet.getRange(lr + 1, 1);
  let missingDomainCursor = { range: missingOutputRange, rowOffset: 0 };

  // Set missing customer id output cursor
  sheet = SpreadsheetApp.getActive().getSheetByName(MISSING_CUSTOMERS);
  lr = sheet.getLastRow();
  let missingCustomerRange = sheet.getRange(lr + 1, 1);
  let missingCustomerCursor = { range: missingCustomerRange, rowOffset: 0 };

  // Set lead output cursor
  sheet = SpreadsheetApp.getActive().getSheetByName(POTENTIAL_LEADS);
  lr = sheet.getLastRow();
  let leadOutputRange = sheet.getRange(lr + 1, 1);
  let leadCursor = { range: leadOutputRange, rowOffset: 0 };

  // Set in play customer cursor
  sheet = SpreadsheetApp.getActive().getSheetByName(IN_PLAY_CUSTOMERS);
  lr = sheet.getLastRow();
  let inPlayCustomerRange = sheet.getRange(lr + 1, 1);
  let inPlayCustomerCursor = { range: inPlayCustomerRange, rowOffset: 0 };

  // Set in play partner cursor
  sheet = SpreadsheetApp.getActive().getSheetByName(IN_PLAY_PARTNERS);
  lr = sheet.getLastRow();
  let inPlayPartnerRange = sheet.getRange(lr + 1, 1);
  let inPlayPartnerCursor = { range: inPlayPartnerRange, rowOffset: 0 };

  // 
  // Step 1: Look for and record email domains that don't have an In Region Customer or Partner account.
  // Record those domains, while also tracking the account id for the "in play" customers and partners.
  //

  let detectedDomains = {}; // For the new stuff
  let loggedLeads = {}; // For the new stuff
  let loggedCustomers = {};
  let loggedPartners = {};
  for (var j = 0; j < inviteInfo.length; j++) {

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

    var attendeeInfo = lookForAccounts_(attendees, customerMap, partnerMap);
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

        if (detectedDomains[d]) {
          continue; // old news
        }

        detectedDomains[d] = true;
        missingDomainCursor.range.offset(missingDomainCursor.rowOffset, FILTER_EMAIL_DOMAIN).setValue(d);     // d is a domain
        missingDomainCursor.rowOffset++;
      }
      for (let j = 0; j < attendees.length; j++) {

        let attendeeEmail = attendees[j];

        // By definition, only emails for domains not yet associated with an account are in the attendee list.
        // However, there may be some hashicorp folks. Filter them out.
        // Record all of these non-hashi emails as potential leads (if an accociated account isn't in Salesforce yet)

        if (loggedLeads[attendeeEmail]) {
          continue;
        }
        if (attendeeEmail.indexOf("hashicorp") != -1) continue;

        loggedLeads[attendeeEmail] = true;
        leadCursor.range.offset(leadCursor.rowOffset, FILTER_EMAIL_DOMAIN).setValue(attendeeEmail);
        leadCursor.rowOffset++;
      }
    }
    else {
      for (c in attendeeInfo.customers) {
        if (loggedCustomers[c]) continue;

        loggedCustomers[c] = true;  // Must use short id! Will be used to filter opportunities. 
        inPlayCustomerCursor.range.offset(inPlayCustomerCursor.rowOffset, FILTER_ACCOUNT_ID).setValue(c);      // c is an Id
        inPlayCustomerCursor.rowOffset++;


      }
      for (p in attendeeInfo.partners) {
        if (loggedPartners[p]) continue;

        loggedPartners[p] = true;
        inPlayPartnerCursor.range.offset(inPlayPartnerCursor.rowOffset, FILTER_ACCOUNT_ID).setValue(p);      // p is an id
        inPlayPartnerCursor.rowOffset++;


      }
    }
  }

  // 
  // Step 2: look for the accounts associated with the missing domains 
  //

  sheet = SpreadsheetApp.getActive().getSheetByName(ALL_CUSTOMERS);
  let rangeData = sheet.getDataRange();
  let lc = rangeData.getLastColumn();
  lr = rangeData.getLastRow();
  if (lr > 1) {

    // This saves info off in emailToCustomerMapG and also updates accountType object. We'll use accoutTypes of EXTERNAL_CUSTOMER_TYPE to persist that info
    load_account_info_chunks_(sheet, lr, lc, detectedDomains, FILTER_TYPE_DOMAIN, EXTERNAL_CUSTOMER_TYPE, CUSTOMER_NAME, CUSTOMER_ID, CUSTOMER_EMAIL_DOMAIN, CUSTOMER_ALT_EMAIL_DOMAINS, customerMap, typeMap, 1, false);
    for (account in typeMap) {

      if (typeMap[account] == EXTERNAL_CUSTOMER_TYPE) {
        missingCustomerCursor.range.offset(missingCustomerCursor.rowOffset, 0).setValue(account);
        missingCustomerCursor.rowOffset++;
        inPlayCustomerCursor.range.offset(inPlayCustomerCursor.rowOffset, FILTER_ACCOUNT_ID).setValue(account);
        inPlayCustomerCursor.rowOffset++;
      }
    }
  }

  //
  // Step 3: add in account overrides to ensure the filters account for them
  //

  let parms = SpreadsheetApp.getActive().getSheetByName(RUN_PARMS);
  let overrideRange = parms.getRange(OVERRIDE_FRAME_ROW, OVERRIDE_FRAME_COL, OVERRIDE_FRAME_ROWS, OVERRIDE_FRAME_COLS); 
  let overrides = overrideRange.getValues();
  for (let row in overrides) {
    if (overrides[row][1]) {
      let overrideAccountId = overrides[row][1];
      if (typeMap[overrideAccountId]) {
        if (typeMap[overrideAccountId] == INTERNAL_CUSTOMER_TYPE) {
          if (!loggedCustomers[overrideAccountId]) {

            loggedCustomers[overrideAccountId] = true;  // Must use short id! Will be used to filter opportunities. 
            inPlayCustomerCursor.range.offset(inPlayCustomerCursor.rowOffset, FILTER_ACCOUNT_ID).setValue(overrideAccountId);
            inPlayCustomerCursor.rowOffset++;
          }
        }
        else if (typeMap[overrideAccountId] == PARTNER_TYPE) {
          if (!loggedPartners[overrideAccountId]) {

            loggedPartners[overrideAccountId] = true;
            inPlayPartnerCursor.range.offset(inPlayPartnerCursor.rowOffset, FILTER_ACCOUNT_ID).setValue(overrideAccountId);
            inPlayPartnerCursor.rowOffset++;
          }
        }
        if (typeMap[overrideAccountId] == EXTERNAL_CUSTOMER_TYPE) {
          if (!loggedCustomers[overrideAccountId]) {

            loggedCustomers[overrideAccountId] = true;  // Must use short id! Will be used to filter opportunities. 
            missingCustomerCursor.range.offset(missingCustomerCursor.rowOffset, 0).setValue(overrideAccountId);
            missingCustomerCursor.rowOffset++;
            inPlayCustomerCursor.range.offset(inPlayCustomerCursor.rowOffset, FILTER_ACCOUNT_ID).setValue(overrideAccountId);
            inPlayCustomerCursor.rowOffset++;
          }
        }
      }
    }
  }

  logOneCol("Identified " + inPlayCustomerCursor.rowOffset + " customers in the current calendar invite set.");
  logOneCol("Identified " + inPlayPartnerCursor.rowOffset + " partners in the current calendar invite set.");
  logOneCol("Identified " + missingCustomerCursor.rowOffset + " potential out-of-region customers.");
  logOneCol("Identified " + leadCursor.rowOffset + " potential lead emails.");


  // Wake up the garbage collection? (this isn't necessary, but why not)
  partnerMap = null;
  customerMap = null;
  typeMap = null;

  logOneCol("End time: " + new Date().toLocaleTimeString());

  // Salesforce OAUTH2 doesn't work, so we have to use Zapier
  // var html = HtmlService.createTemplateFromFile('index').evaluate().setWidth(500);
  // SpreadsheetApp.getUi().showSidebar(html);
}

function createChoiceLists() {


  clearTab_(CHOICE_ACCOUNT, [['Account Name', 'Account Id']]);
  clearTab_(CHOICE_PARTNER, [['Account Name', 'Account Id']]);
  clearTab_(CHOICE_OP, [['Op Name', 'Op Id']]);
  clearTab_(CHOICE_LEAD, [['Lead Email']]);

  //
  // Build the choice lists for the Review tab. We need to get all the duplicate accounts, not just the primary, in case we selected the wrong one.
  //   

  logOneCol("Building choice lists " + new Date().toLocaleTimeString());

  // Set customer choices cursor
  sheet = SpreadsheetApp.getActive().getSheetByName(CHOICE_ACCOUNT);
  lr = sheet.getLastRow();
  let choiceCustomerRange = sheet.getRange(lr + 1, 1);
  let choiceCustomerCursor = { range: choiceCustomerRange, rowOffset: 0 };

  // Set partner choices cursor
  sheet = SpreadsheetApp.getActive().getSheetByName(CHOICE_PARTNER);
  lr = sheet.getLastRow();
  let choicePartnerRange = sheet.getRange(lr + 1, 1);
  let choicePartnerCursor = { range: choicePartnerRange, rowOffset: 0 };

  // Set op choices cursor
  sheet = SpreadsheetApp.getActive().getSheetByName(CHOICE_OP);
  lr = sheet.getLastRow();
  let choiceOpRange = sheet.getRange(lr + 1, 1);
  let choiceOpCursor = { range: choiceOpRange, rowOffset: 0 };


  // Set lead choices cursor
  sheet = SpreadsheetApp.getActive().getSheetByName(CHOICE_LEAD);
  lr = sheet.getLastRow();
  let choiceLeadRange = sheet.getRange(lr + 1, 1);
  let choiceLeadCursor = { range: choiceLeadRange, rowOffset: 0 };

  let opFilter = {}

  for (cn in primaryAccountsG) {
    // First part of choice list init. See finishChoiceLists for second part.
    let id = primaryAccountsG[cn];
    choiceCustomerCursor.range.offset(choiceCustomerCursor.rowOffset, CHOICE_ACCOUNT_NAME).setValue(cn);
    choiceCustomerCursor.range.offset(choiceCustomerCursor.rowOffset, CHOICE_ACCOUNT_ID).setValue(id);
    choiceCustomerCursor.rowOffset++;
    opFilter[id.substring(0, id.length - 3)] = true;
  }
  for (pn in primaryPartnersG) {
    // First part of choice list init. See finishChoiceLists for second part.
    let id = primaryPartnersG[pn];
    choicePartnerCursor.range.offset(choicePartnerCursor.rowOffset, CHOICE_PARTNER_NAME).setValue(pn);
    choicePartnerCursor.range.offset(choicePartnerCursor.rowOffset, CHOICE_PARTNER_ID).setValue(id);
    choicePartnerCursor.rowOffset++;
  }
  for (cn in duplicateAccountsG) {
    let id = duplicateAccountsG[cn];
    choiceCustomerCursor.range.offset(choiceCustomerCursor.rowOffset, CHOICE_ACCOUNT_NAME).setValue(cn);
    choiceCustomerCursor.range.offset(choiceCustomerCursor.rowOffset, CHOICE_ACCOUNT_ID).setValue(id);
    choiceCustomerCursor.rowOffset++;
    opFilter[id.substring(0, id.length - 3)] = true;
  }
  for (pn in duplicatePartnersG) {
    let id = duplicatePartnersG[pn];
    choicePartnerCursor.range.offset(choicePartnerCursor.rowOffset, CHOICE_PARTNER_NAME).setValue(pn);
    choicePartnerCursor.range.offset(choicePartnerCursor.rowOffset, CHOICE_PARTNER_ID).setValue(id);
    choicePartnerCursor.rowOffset++;
  }

  //
  // Ops
  //

  sheet = SpreadsheetApp.getActive().getSheetByName(OPPORTUNITIES);
  rangeData = sheet.getDataRange();
  let olc = rangeData.getLastColumn();
  let olr = rangeData.getLastRow();

  if (olc >= OP_COLUMNS) {
    let chunkSize = 1000;
    let chunkFirstRow = 2;
    let chunkLastRow = chunkSize < olr ? chunkSize : olr;
    let chunkLength = chunkLastRow - chunkFirstRow + 1;
    while (chunkLength > 0) {
      let scanRange = sheet.getRange(chunkFirstRow, 1, chunkLength, olc);
      let opInfo = scanRange.getValues();

      for (let j = 0; j < chunkLength; j++) {

        if (!opFilter[opInfo[j][OP_ACCOUNT_ID]]) {
          continue;
        }
        choiceOpCursor.range.offset(choiceOpCursor.rowOffset, CHOICE_OP_NAME).setValue(opInfo[j][OP_NAME]);
        choiceOpCursor.range.offset(choiceOpCursor.rowOffset, CHOICE_OP_ID).setValue(opInfo[j][OP_ID]);
        choiceOpCursor.rowOffset++;
      }

      chunkFirstRow = chunkLastRow + 1;
      chunkLastRow += chunkSize;
      chunkLength = chunkSize;
      if (chunkLastRow > olr) {
        chunkLastRow = olr;
        chunkLength = chunkLastRow - chunkFirstRow + 1
      }
    }
  }

  //
  // Leads - these are just the "real leads" from all the potential lead emails. 
  // TODO - provide other real leads for the domain which has no account record. 
  //

  for (em in emailToLeadMapG) {
    choiceLeadCursor.range.offset(choiceLeadCursor.rowOffset, CHOICE_LEAD_EMAIL).setValue(em);
    choiceLeadCursor.rowOffset++;
  }


  logOneCol("Choice list build complete " + new Date().toLocaleTimeString());
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
    else Logger.log("WARNING : " + primary + "is not a known primary product!");
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
    let range = sheet.getRange(1, 1, 1, header[0].length);
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
  let scanRange = sheet.getRange(2, 1, elr - 1, elc); // 2 is to skip header. Assumes a header!!!!
  let eventInfo = scanRange.getValues();

  if (elc < EVENT_COLUMNS) {
    Logger.log("ERROR: Imported Customers was only " + elc + " fields wide. Not enough! Something is wrong.");
    return;
  }

  // Set event output cursor
  let sheet2 = SpreadsheetApp.getActive().getSheetByName(UPLOAD_STAGE);
  let ulr = sheet2.getLastRow();
  let uploadRange = sheet2.getRange(ulr + 1, 1);
  let rowOffset = 0;
  let uc = 0;


  for (j = 0; j < elr - 1; j++) {

    if (eventInfo[j][EVENT_PROCESS] == PROCESS_SKIP) {
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
    uploadRange.offset(rowOffset, EVENT_NOTES).setValue(eventInfo[j][EVENT_NOTES]);
    uploadRange.offset(rowOffset, EVENT_LEAD).setValue(eventInfo[j][EVENT_LEAD]);
    uc++;
    rowOffset++;
  }
  logOneCol("Uploaded " + uc + " events.");

}

function load_account_info_chunks_(accountInfoSheet, totalNumberOfRows, numberOfColumns, filter, filterType, accountType, accountNameIdx, accountIdIdx, emailIdx, altEmailIdx, emailToAccountMap, accountTypeMap, phase, loggingEnabled) {

  let chunkSize = 1000;
  let chunkFirstRow = 2;
  let chunkLastRow = chunkSize < totalNumberOfRows ? chunkSize : totalNumberOfRows;
  let chunkLength = chunkLastRow - chunkFirstRow + 1;
  while (chunkLength > 0) {
    let scanRange = accountInfoSheet.getRange(chunkFirstRow, 1, chunkLength, numberOfColumns);
    let accountInfo = scanRange.getValues();


    load_account_info_worker_(accountInfo, chunkLength, filter, filterType, accountType, accountNameIdx, accountIdIdx, emailIdx, altEmailIdx, emailToAccountMap, accountTypeMap, phase, loggingEnabled);

    chunkFirstRow = chunkLastRow + 1;
    chunkLastRow += chunkSize;
    chunkLength = chunkSize;
    if (chunkLastRow > totalNumberOfRows) {
      chunkLastRow = totalNumberOfRows;
      chunkLength = chunkLastRow - chunkFirstRow + 1
    }
  }
}

function load_account_info_worker_(accountInfo, numberOfRows, filter, filterType, accountType, accountNameIdx, accountIdIdx, emailIdx, altEmailIdx, emailToAccountMap, accountTypeMap, phase, loggingEnabled) {
  // accountInfo is an array of account records. numberOfRows is the number of account rounds. 
  // targets is an import filter, an inventory of the domains we are targeting for load, accounts that are not in region 
  // or may not exist yet based on the invite attendees for this run. This is to reduce the amount of information we need to keep in memory.
  // accountType is either in-region-customer, customer (all), partner or lead
  // The <name>Idx parameters are the indexes of the particular field in the accountInfo table.
  // emailToAccountMap is an object log of what email domain belongs to what account (right now it is actually a global variable)
  // loggingEnabled there because we run this in two phases. Only want to publish logs on the second phase to avoid redundant info.
  // One global variable is updated (other than the one passed in as emailToAccountMap), and that is accountType. 
  // It allows us to track the type of each account we are processing.


  // This normally processes email domains (on stuff after the @ sign). However, one of the callers (for leads) will pass in a full email (<x>@<domain>). So, make
  // sure the logic is generic enough to handle that case. 

  let accountLog = {};

  // Build the email domain (or full email for leads) to account mapping  
  for (let j = 0; j < numberOfRows; j++) {

    let id = accountInfo[j][accountIdIdx].trim();

    if (filter && filterType == FILTER_TYPE_ID && !filter[id]) {
      continue;
    }

    let emailDomainString = accountInfo[j][emailIdx];
    let altDomainString = 0;
    if (altEmailIdx) {
      altDomainString = accountInfo[j][altEmailIdx];
    }

    if (!emailDomainString && !altDomainString) {
      if (loggingEnabled && accountType != LEAD_TYPE) logOneCol("NOTICE - " + accountInfo[j][accountNameIdx] + " has no email domain!");
      // Thousands of leads don't have an email. Don't try to log!!! Email domains are critical for accounts however.
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

    for (let k = 0; k < emailDomains.length; k++) {


      // LATAM domains violate rules
      //let emailRegex = /[-\w]+\.[a-zA-Z]{2,3}$/;
      //let domain = emailRegex.exec(emailDomains[k].trim());
      let superDomain = emailDomains[k].trim(); // probably just one domain, but you never know

      // Because sometimes Reps use spaces instead of commas to separate domains in
      // the salesforce field. Deal with it now.
      let potentiallyMoreDomains = superDomain.split(" ");
      for (let i = 0; i < potentiallyMoreDomains.length; i++) {

        let domain = potentiallyMoreDomains[i].trim();

        // FIXME are there other prefixes I don't know about? 
        if (domain.indexOf("www.") == 0) {
          domain = domain.substring(4);
        }

        if (filter && filterType == FILTER_TYPE_DOMAIN && !filter[domain]) {
          continue;
        }

        if (uniqueDomains[domain]) {
          continue; // Rep put this domain in twice!
        }
        uniqueDomains[domain] = true;

        if (accountLog[domain]) {
          //Logger.log("WARNING : Found account with a duplicate email domain - " + accountInfo[j][accountNameIdx] + ":" + domain);


          if (2 == phase) {
            // Don't report bad Salesforce data if invites don't reference it

            logFourCol("NOTICE - Multiple accounts for email domain!", "Domain: " + domain, "Rejected: " + accountInfo[j][accountNameIdx], "Selected: " + accountLog[domain]);

            // Remember the dup (for choice lists)
            // It is the case that an account with multiple domains could be primary on one domain, but not-primary on another.
            // Therefore, it is possible that an account was already put into the other queue already. Don't allow for duplicate accounts.
            switch (accountType) {
              case INTERNAL_CUSTOMER_TYPE:
              case EXTERNAL_CUSTOMER_TYPE:
                if (!primaryAccountsG[accountInfo[j][accountNameIdx]])
                  duplicateAccountsG[accountInfo[j][accountNameIdx]] = id;
                break;
              case PARTNER_TYPE:
                if (!primaryPartnersG[accountInfo[j][accountNameIdx]])
                  duplicatePartnersG[accountInfo[j][accountNameIdx]] = id;
                break;
              default:
                break;
            }
          }
          continue; // Take the first
        }

        accountLog[domain] = accountInfo[j][accountNameIdx];
        emailToAccountMap[domain] = id;
        accountTypeMap[id] = accountType;
        if (2 == phase) {
          // Remember the primary (for choice list)
          // It is the case that an account with multiple domains could be primary on one domain, but not-primary on another.
          // Therefore, it is possible that an account was already put into the other queue already. Don't allow for duplicate accounts.
          switch (accountType) {
            case INTERNAL_CUSTOMER_TYPE:
            case EXTERNAL_CUSTOMER_TYPE:
              if (!duplicateAccountsG[accountInfo[j][accountNameIdx]])
                primaryAccountsG[accountInfo[j][accountNameIdx]] = id;
              break;
            case PARTNER_TYPE:
              if (!duplicatePartnersG[accountInfo[j][accountNameIdx]])
                primaryPartnersG[accountInfo[j][accountNameIdx]] = id;
              break;
            default:
              break;
          }
        }
      }
    }
  }
}

function tester() { // DAK
  logStamp("Tester");
  load_partner_info_(false, emailToPartnerMapG, accountTypeG, true);
}

function load_customer_info_(inRegionOnly, emailToAccountMap, accountTypeMap, phase, loggingEnabled) {
  // Returns false on error 
  // Updates emailToCustomerMapG and accountTypeG

  //
  // Load up our region's customers
  //

  let targetedAccounts = null;
  if (!inRegionOnly) {
    targetedAccounts = loadFilter_(IN_PLAY_CUSTOMERS, FILTER_ACCOUNT_ID, false);
  }

  let sheet = SpreadsheetApp.getActive().getSheetByName(IN_REGION_CUSTOMERS);
  let rangeData = sheet.getDataRange();
  let alc = rangeData.getLastColumn();
  let alr = rangeData.getLastRow();
  if (alr < 2) {
    // No in region customers (that's odd)
    if (loggingEnabled) logOneCol("WARNING - No in-region customers found! Perhaps you should refresh the In Region Customer tab.");
  }
  else {

    let scanRange = sheet.getRange(2, 1, alr - 1, alc);
    let customerInfo = scanRange.getValues();

    if (alc < CUSTOMER_COLUMNS) {
      logOneCol("ERROR: Imported In Region Customers was only " + alc + " fields wide. Not enough! Something is wrong.");
      return false;
    }

    load_account_info_worker_(customerInfo, alr - 1, targetedAccounts, FILTER_TYPE_ID, INTERNAL_CUSTOMER_TYPE, CUSTOMER_NAME, CUSTOMER_ID, CUSTOMER_EMAIL_DOMAIN, CUSTOMER_ALT_EMAIL_DOMAINS, emailToAccountMap, accountTypeMap, phase, loggingEnabled);
  }

  if (inRegionOnly) return true;

  // 
  // Get the target list, accounts suspected to be out-of-region customers, to filter out what we need to save in memory.
  //

  targetedAccounts = loadFilter_(MISSING_CUSTOMERS, FILTER_ACCOUNT_ID, false);

  //
  // Load up customers outside of our region, but only the ones that may be needed (in the targeted list)
  //

  sheet = SpreadsheetApp.getActive().getSheetByName(ALL_CUSTOMERS);
  rangeData = sheet.getDataRange();
  alc = rangeData.getLastColumn();
  alr = rangeData.getLastRow();
  if (alr < 2) {
    // No external customers logged
    if (loggingEnabled) logOneCol("INFO - No out-region customers found! Perhaps you should run Activity > Import Missing Accounts.");
    return true;
  }

  // This list could get massive. Do it in baby steps
  load_account_info_chunks_(sheet, alr, alc, targetedAccounts, FILTER_TYPE_ID, EXTERNAL_CUSTOMER_TYPE, CUSTOMER_NAME, CUSTOMER_ID, CUSTOMER_EMAIL_DOMAIN, CUSTOMER_ALT_EMAIL_DOMAINS, emailToAccountMap, accountTypeMap, phase, loggingEnabled);

  return true;
}

function load_partner_info_(allPartners, emailToAccountMap, accountTypeMap, phase, loggingEnabled) {

  // Updates emailToPartnerMapG and accountTypeG

  let targets = null;
  if (!allPartners) {
    targets = loadFilter_(IN_PLAY_PARTNERS, FILTER_ACCOUNT_ID, false);
  }

  let sheet = SpreadsheetApp.getActive().getSheetByName(PARTNERS);
  let rangeData = sheet.getDataRange();
  let plc = rangeData.getLastColumn();
  let plr = rangeData.getLastRow();
  if (plr < 2) {
    // No partners (that's odd)
    logOneCol("WARNING : No Partners found! Perhaps you should refresh the partner tab.");
  }
  else {
    load_account_info_chunks_(sheet, plr, plc, targets, FILTER_TYPE_ID, PARTNER_TYPE, PARTNER_NAME, PARTNER_ID, PARTNER_EMAIL_DOMAIN, PARTNER_ALT_EMAIL_DOMAINS, emailToAccountMap, accountTypeMap, phase, loggingEnabled);
  }
  return true;
}

function load_lead_info_(loggingEnabled) {

  // 
  // Get the target list, accounts suspected to only have leads, to filter out what we need to save in memory.
  // Updates emailToLeadMapG and accountTypeG
  //

  let targetedLeads = loadFilter_(POTENTIAL_LEADS, FILTER_EMAIL_DOMAIN, false);

  //
  // Import the targeted lead info
  //

  let sheet = SpreadsheetApp.getActive().getSheetByName(LEADS);
  let rangeData = sheet.getDataRange();
  let lastColumn = rangeData.getLastColumn();
  let lastRow = rangeData.getLastRow();

  if (lastColumn < LEAD_COLUMNS) {
    Logger.log("ERROR: " + LEADS + " does not have " + LEAD_COLUMNS + " fields.");
    return;
  }

  load_account_info_chunks_(sheet, lastRow, lastColumn, targetedLeads, FILTER_TYPE_DOMAIN, LEAD_TYPE, LEAD_NAME, LEAD_ID, LEAD_EMAIL, 0, emailToLeadMapG, accountTypeG, 2, loggingEnabled)
}

function loadFilter_(tabName, fieldNumber, truncate) {
  // Create a simple "filter map" keyed by the value in the fieldNumber field of this tabName. The map value will be set to true;
  // Assumes the table has a header in row 1! (Skips to row 2 for data start)
  // 
  // What's the point of the truncate parm? Almost everything refers to objects in Salesforce by there 18 digit "long id", except ...
  // Opportunties. They use the shorter 15 digit id to reference an account. 
  // So, when filtering a list of opportunities, we need to use a "truncated" list (15 digit account ids). 
  // So, truncate what will be an 18 digit id in fieldNumber down to 15.

  let targets = {};

  try {
    let sheet = SpreadsheetApp.getActive().getSheetByName(tabName);
    let rangeData = sheet.getDataRange();
    let lc = rangeData.getLastColumn();
    let lr = rangeData.getLastRow();
    if (lr < 2) {
      // No external customers in this run
      return targets; // Still need to return the empty filter (its really a mask, so filter out everything)
    }
    let scanRange = sheet.getRange(2, 1, lr - 1, lc);
    let v = scanRange.getValues();
    for (j = 0; j < lr - 1; j++) {
      if (truncate) {
        let k = v[j][fieldNumber]
        targets[k.substring(0, k.length - 3)] = true;
      }
      else {
        targets[v[j][fieldNumber]] = true;
      }
    }
  }
  catch (e) {
    logOneCol("ERROR - loadFilter__ issue on " + tabName + ": " + e);
  }
  return targets;
}

function printIt_() {
  let keys = PropertiesService.getScriptProperties().getKeys();
  for (var i = 0; i < keys.length; i++) {
    Logger.log(keys[i] + ":" + PropertiesService.getScriptProperties().getProperty(keys[i]));
  }
}

function findLead_(attendees) {

  for (let i = 0; i < attendees.length; i++) {
    if (emailToLeadMapG[attendees[i]]) {
      return emailToLeadMapG[attendees[i]];
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
      let rowCount = lastRow - fromRow + 1
      let scanRange = sheet.getRange(fromRow, 1, rowCount, lastColumn);
      let inputArray = scanRange.getValues();
      return inputArray;
    }
  }
  catch (e) {
    logOneCol("ERROR - load_tab_ issue on " + sheetName + ": " + e);
  }

  return [];
}

function load_map_(tabName, firstRow, tabColumns, keyField, valueField, filter, filterField, existingMap) {
  let map = {};
  if (existingMap) map = existingMap;

  try {
    let sheet = SpreadsheetApp.getActive().getSheetByName(tabName);
    let rangeData = sheet.getDataRange();
    let lastColumn = rangeData.getLastColumn();
    let lastRow = rangeData.getLastRow();

    if (lastColumn < tabColumns) {
      Logger.log("ERROR: " + tabName + " does not have " + tabColumns + " fields.");
      return map;
    }

    let chunkSize = 1000;
    let chunkFirstRow = firstRow;
    let chunkLastRow = chunkSize < lastRow ? chunkSize : lastRow;
    let chunkLength = chunkLastRow - chunkFirstRow + 1;
    while (chunkLength > 0) {
      let scanRange = sheet.getRange(chunkFirstRow, 1, chunkLength, lastColumn);
      let info = scanRange.getValues();

      for (let j = 0; j < chunkLength; j++) {

        if (filter && !filter[info[j][filterField]]) {
          continue;
        }
        map[info[j][keyField]] = info[j][valueField];

      }

      chunkFirstRow = chunkLastRow + 1;
      chunkLastRow += chunkSize;
      chunkLength = chunkSize;
      if (chunkLastRow > lastRow) {
        chunkLastRow = lastRow;
        chunkLength = chunkLastRow - chunkFirstRow + 1
      }
    }
  }
  catch (e) {
    logOneCol("ERROR - load_map_ issue on " + tabName + ": " + e);
  }

  return map;
}

function analyzeSubject_(text) {

  let rv = { prepTime: 0 }

  let subject = text.trim();
  if (prepCalendarEntries[subject]) {
    rv.prepTime = prepCalendarEntries[subject];
    prepCalendarEntries[subject] = 0;
  }
  return rv;
}

function generateLongId_(id) {

  // Transform a 15 character case sensitive Salesforce ID into the 
  // "long version", an 18 character case insensitive id.
  // This is Salesforce's algorithm.

  let retVal = id;

  if (id.length == 15) {

    let addon = "";
    for (let block = 0; block < 3; block++) {
      let loop = 0;
      for (let position = 0; position < 5; position++) {
        let current = id.charAt(block * 5 + position);
        if (current >= "A" && current <= "Z")
          loop += 1 << position;
      }
      addon += "ABCDEFGHIJKLMNOPQRSTUVWXYZ012345".charAt(loop);
    }
    retVal = (id + addon);
  }
  return retVal;
}

function incrementAgendaItemCounts_(ledger, eventInfo) {
  // This must be keep in sync with the agendaItemsToTrack object in eheuristic.gs

  if (eventInfo[EVENT_DEMO]) {
    ledger.demo++;
  }
  if (eventInfo[EVENT_POV]) {
    ledger.pov++;
  }
  if (eventInfo[EVENT_WORKSHOP]) {
    ledger.workshop++;
  }
  if (eventInfo[EVENT_DIVE]) {
    ledger.deepdive++;
  }
}

function collectStats_(assignedTo, eventType, eventInfo, relatedTo, attendeeString) {

  if (!statsOutputWorksheetUrlG) {
    // If we have no place to send it, don't bother collecting it
    return;
  }

  // attendees is a comma separated string of meeting attendees.
  // relatedTo will be null for leads

  //let sId = staffEmailToIdMapG[inviteInfo[ASSIGNED_TO]];

  // statsLedgerG Structure
  //
  // ledger
  //  global
  //   agendaItems
  //    demo : count
  //    pov : count
  //    workshop : count
  //    deepdive : count
  //  users
  //   userId ...
  //    generalCounts
  //      op : count
  //      op_new : count
  //      customer : count
  //      customer_new : count
  //      partner : count
  //      partner_new : count
  //      lead : count 
  //      lead_new : count
  //.     adendaItems
  //        demo : count
  //        pov : count
  //        workshop : count
  //        deepdive : count
  //    ops
  //      opId ...
  //        new_count : count
  //        recurring_count : count
  //        agendaItems
  //          demo : count
  //          pov : count
  //          workshop : count
  //          deepdive : count
  //        attendees
  //          attendee ...
  //            new_count : count
  //            recurring_count : count
  //            agendaItems
  //              demo : count
  //              pov : count
  //              workshop : count
  //              deepdive : count
  //    customers
  //      accountId ...
  //        new_count : count
  //        recurring_count : count
  //        agendaItems
  //          demo : count
  //          pov : count
  //          workshop : count
  //          deepdive : count
  //        attendees
  //         attendee ...
  //          new_count : count
  //          recurring_count : count
  //          agendaItems
  //              demo : count
  //              pov : count
  //              workshop : count
  //              deepdive : count
  //    partners
  //      accountId ...
  //        new_count : count
  //        recurring_count : count
  //        adendaItems
  //          demo : count
  //          pov : count
  //          workshop : count
  //          deepdive : count
  //        attendees
  //         attendee ...
  //           new_count : count
  //           recurring_count : count
  //           agendaItems
  //              demo : count
  //              pov : count
  //              workshop : count
  //              deepdive : count
  //    leads
  //      attendee : ...
  //           new_count : count
  //           recurring_count : count


  let attendees = attendeeString.split(",");

  if (!statsLedgerG.users[assignedTo]) {
    let agendaItems = {demo : 0, pov : 0, workshop : 0, deepdive : 0 };
    let meetingCounts = { op: 0, op_new: 0, customer: 0, customer_new : 0, partner: 0, partner_new: 0, lead: 0, lead_new : 0, agendaItems: agendaItems};
    statsLedgerG.users[assignedTo] = { generalCounts: meetingCounts, ops: {}, customers: {}, partners: {}, leads: {} };
    // logOneCol("Collecting contact stats for " + assignedTo);
  }

  incrementAgendaItemCounts_(statsLedgerG.global.agendaItems, eventInfo);
  incrementAgendaItemCounts_(statsLedgerG.users[assignedTo].generalCounts.agendaItems, eventInfo);

  switch (eventType) {
    case OP_EVENT:
      statsLedgerG.users[assignedTo].generalCounts.op++;
      if (!statsLedgerG.users[assignedTo].ops[relatedTo]) {
        let agendaItems = {demo : 0, pov : 0, workshop : 0, deepdive : 0 };
        statsLedgerG.users[assignedTo].ops[relatedTo] = { new_count: 0, recurring_count: 0, agendaItems: agendaItems, attendees: {} };
      }
      incrementAgendaItemCounts_(statsLedgerG.users[assignedTo].ops[relatedTo].agendaItems, eventInfo);
      if (eventInfo[EVENT_RECURRING]) {
        statsLedgerG.users[assignedTo].ops[relatedTo].recurring_count++;
      }
      else {
        statsLedgerG.users[assignedTo].generalCounts.op_new++;
        statsLedgerG.users[assignedTo].ops[relatedTo].new_count++;
      }
      break;
    case CUSTOMER_EVENT:
      statsLedgerG.users[assignedTo].generalCounts.customer++;
      if (!statsLedgerG.users[assignedTo].customers[relatedTo]) {
        let agendaItems = {demo : 0, pov : 0, workshop : 0, deepdive : 0 };
        statsLedgerG.users[assignedTo].customers[relatedTo] = { new_count: 0, recurring_count: 0, agendaItems: agendaItems, attendees: {} }; // Object for all possible attendee counts
      }
      incrementAgendaItemCounts_(statsLedgerG.users[assignedTo].customers[relatedTo].agendaItems, eventInfo);
      if (eventInfo[EVENT_RECURRING]) {
        statsLedgerG.users[assignedTo].customers[relatedTo].recurring_count++;
      }
      else {
        statsLedgerG.users[assignedTo].generalCounts.customer_new++;
        statsLedgerG.users[assignedTo].customers[relatedTo].new_count++;
      }
      break;
    case PARTNER_EVENT:
      statsLedgerG.users[assignedTo].generalCounts.partner++;
      if (!statsLedgerG.users[assignedTo].partners[relatedTo]) {
        let agendaItems = {demo : 0, pov : 0, workshop : 0, deepdive : 0 };
        statsLedgerG.users[assignedTo].partners[relatedTo] = { new_count: 0, recurring_count: 0, agendaItems: agendaItems, attendees: {} }; // Object for all possible attendee counts
      }
      incrementAgendaItemCounts_(statsLedgerG.users[assignedTo].partners[relatedTo].agendaItems, eventInfo);
      if (eventInfo[EVENT_RECURRING]) {
        statsLedgerG.users[assignedTo].partners[relatedTo].recurring_count++;
      }
      else {
        statsLedgerG.users[assignedTo].generalCounts.partner_new++;
        statsLedgerG.users[assignedTo].partners[relatedTo].new_count++;
      }
      break;
    case LEAD_EVENT:
      statsLedgerG.users[assignedTo].generalCounts.lead++;
      if (!eventInfo[EVENT_RECURRING]) {
        statsLedgerG.users[assignedTo].generalCounts.lead_new++;
      }
      break;
    default:
      break;
  }

  for (let i = 0; i < attendees.length; i++) {
    if (attendees[i].indexOf("hashicorp.com") == -1) {
      switch (eventType) {
        case OP_EVENT:       
          incrementStat_(statsLedgerG.users[assignedTo].ops[relatedTo].attendees, attendees[i], eventInfo);
          break;
        case CUSTOMER_EVENT:         
          incrementStat_(statsLedgerG.users[assignedTo].customers[relatedTo].attendees, attendees[i], eventInfo);
          break;
        case PARTNER_EVENT:       
          incrementStat_(statsLedgerG.users[assignedTo].partners[relatedTo].attendees, attendees[i], eventInfo);
          break;
        case LEAD_EVENT:
          incrementStat_(statsLedgerG.users[assignedTo].leads, attendees[i], eventInfo);
          break;
        default:
          break;
      }
    }
  }
}

function loadMeetingMap_(map, propertiesTab) {

  // We do not want people using their own properties because we want consistent results, so
  // I'm going to hardcode a bunch of stuff in here to at least make it a little difficult to change
  // stuff.

  // Assets there is a map to load.
  // Initialize the map

  let activityMachineProperties = "https://docs.google.com/spreadsheets/d/1BrNyxuUdl9aKNwGfAwtcQbGckOA4YUMrfyojGhAiIos/edit#gid=602540278";
  let worksheet = SpreadsheetApp.openByUrl(activityMachineProperties);
  if (!worksheet) {
    logOneCol("ERROR - Activity Machine's organization wide properties not found!");
    return false;
  }

  let sheet = worksheet.getSheetByName(propertiesTab);

  if (!sheet) {
    logOneCol("Error - could not find the " + propertiesTab + " tab in the Activity Machine properties");
    return false;
  }

  let rangeData = sheet.getDataRange();
  let lc = rangeData.getLastColumn();
  let lr = rangeData.getLastRow();
  if (lr < 2) {
    // No in region customers (that's odd)
    logOneCol("ERROR - Found an empty " + propertiesTab);
    return false;
  }

  if (lc < 4) {
    logOneCol("ERROR: " + propertiesTab + " is missing some columns.");
    return false;
  }

  let mapRange = sheet.getRange(2, 1, lr - 1, lc);
  let mappings = mapRange.getValues();

  for (i = 0; i < lr - 1; i++) {
    let m = { key: mappings[i][0], meeting: mappings[i][1], regex: new RegExp(mappings[i][2]), stage: mappings[i][3] };
    map.push(m);
  }
  return true;
}

function printStats_() {

  // Dump stats
  // See collectStats method for statsLedgerG schema
  //

  if (!statsOutputWorksheetUrlG) {
    return;
  }

  let parms = SpreadsheetApp.getActive().getSheetByName(RUN_PARMS);
  let dateRange = parms.getRange(DATE_FRAME_ROW, DATE_FRAME_COL, 2, 1); 
  let dates = dateRange.getValues();

  let statsOutputIdRange = parms.getRange(STATS_FRAME_ROW, STATS_FRAME_COL); 
  let rangeValues = statsOutputIdRange.getValues();
  statsOutputWorksheetUrlG = rangeValues[0][0];

  let chartLedger = [];
  let chartIndex = 0;
  let minChartSpacing = 21; // Need 21 Rows to fit chart. Small datasets need to be spaced at least 21 rows

  let worksheet = SpreadsheetApp.openByUrl(statsOutputWorksheetUrlG);
  if (!worksheet) {
    logOneCol("ERROR - Stats output worksheet URL is not valid: " + statsOutputWorksheetUrlG);
    return;
  }

  let sheet = worksheet.getSheetByName(MEETING);
  //let sheet = SpreadsheetApp.getActive().getSheetByName(STATS);
  if (!sheet) {
    sheet = worksheet.getSheets()[0];
    if (!sheet) {
      logOneCol("Error - could not find a tab for stats output in " + worksheet.getName());
      return;
    }
  }
  let tabName = sheet.getName();

  logOneCol("Logging stats to " + tabName + " in "+ worksheet.getName());
  sheet.clearContents();
  deleteCharts_(sheet);

  let outputRange = sheet.getRange(MEETING_COUNTS_ORIGIN_ROW, MEETING_COUNTS_ORIGIN_COL);
  let rowOffset = 0;

  let from = Utilities.formatDate(new Date(dates[0][0]), "GMT", "MM/dd/yyyy");
  let to = Utilities.formatDate(new Date(dates[1][0]), "GMT", "MM/dd/yyyy");
  let when = new Date().toLocaleDateString() + " " + new Date().toLocaleTimeString()
  outputRange.offset(rowOffset++, 0).setValue("#################################################");
  outputRange.offset(rowOffset++, 0).setValue("## ----- Event Statistics from " + from + " to " + to + " ----- ##");
  outputRange.offset(rowOffset++, 0).setValue("## ----- Run executed at " + when + " --------------- ##");
  outputRange.offset(rowOffset++, 0).setValue("#################################################");

  rowOffset = rowOffset + 3;
  outputRange.offset(rowOffset++, 0).setValue("------------------------------------------------------------------------");
  outputRange.offset(rowOffset++, 0).setValue("                               Group Data");
  outputRange.offset(rowOffset++, 0).setValue("------------------------------------------------------------------------");
  rowOffset++;

  // Calculate statistics
  let totalMeetings = 0;
  let totalNewMeetings = 0;
  let userCount = 0;
  let statistics = {};
  for (user in statsLedgerG.users) {
    userCount++;
    statistics[user] = {ops : 0, customers : 0, partners : 0, new_meetings : 0, recurring_meetings : 0, attendees_new : 0, attendees_recurring : 0 };
    totalMeetings = totalMeetings + statsLedgerG.users[user].generalCounts.op + statsLedgerG.users[user].generalCounts.customer +
      statsLedgerG.users[user].generalCounts.partner + statsLedgerG.users[user].generalCounts.lead;
    totalNewMeetings = totalNewMeetings + statsLedgerG.users[user].generalCounts.op_new + statsLedgerG.users[user].generalCounts.customer_new +
      statsLedgerG.users[user].generalCounts.partner_new + statsLedgerG.users[user].generalCounts.lead_new;

    for (op in statsLedgerG.users[user].ops) {
      statistics[user].ops++;
      statistics[user].new_meetings = statistics[user].new_meetings + statsLedgerG.users[user].ops[op].new_count;
      statistics[user].recurring_meetings = statistics[user].recurring_meetings + statsLedgerG.users[user].ops[op].recurring_count;
      for (d in statsLedgerG.users[user].ops[op].attendees) {
        if (statsLedgerG.users[user].ops[op].attendees[d].new_count > 0) {
          statistics[user].attendees_new++;
        }
        if (statsLedgerG.users[user].ops[op].attendees[d].recurring_count > 0) {
          statistics[user].attendees_recurring++;
        }
      }
    }

    for (account in statsLedgerG.users[user].customers) {
      statistics[user].customers++;
      statistics[user].new_meetings = statistics[user].new_meetings + statsLedgerG.users[user].customers[account].new_count;
      statistics[user].recurring_meetings = statistics[user].recurring_meetings + statsLedgerG.users[user].customers[account].recurring_count;
      for (d in statsLedgerG.users[user].customers[account].attendees) {
        if (statsLedgerG.users[user].customers[account].attendees[d].new_count > 0) {
          statistics[user].attendees_new++;
        }
        if (statsLedgerG.users[user].customers[account].attendees[d].recurring_count > 0) {
          statistics[user].attendees_recurring++;
        }
      }
    }

    for (account in statsLedgerG.users[user].partners) {
      statistics[user].partners++;
      statistics[user].new_meetings = statistics[user].new_meetings + statsLedgerG.users[user].partners[account].new_count;
      statistics[user].recurring_meetings = statistics[user].recurring_meetings + statsLedgerG.users[user].partners[account].recurring_count;
      for (d in statsLedgerG.users[user].partners[account].attendees) {
        if (statsLedgerG.users[user].partners[account].attendees[d].new_count > 0) {
          statistics[user].attendees_new++;
        }
        if (statsLedgerG.users[user].partners[account].attendees[d].recurring_count > 0) {
          statistics[user].attendees_recurring++;
        }
      }
    }

    for (lead in statsLedgerG.users[user].leads) {
      // FIXME dak
      statistics[user].new_meetings = statistics[user].new_meetings + statsLedgerG.users[user].leads[lead].new_count;
      statistics[user].recurring_meetings = statistics[user].recurring_meetings + statsLedgerG.users[user].leads[lead].recurring_count;

      if (statsLedgerG.users[user].leads[lead].new_count > 0) {
        statistics[user].attendees_new++;
      }
      if (statsLedgerG.users[user].leads[lead].recurring_count > 0) {
        statistics[user].attendees_recurring++;
      }

    }

  }

  if (userCount == 0) return;

  outputRange.offset(rowOffset++, 0).setValue("----------------- Statistics -----------------");
  outputRange.offset(rowOffset++, 0).setValue("Total meeting count: " + totalMeetings);
  outputRange.offset(rowOffset++, 0).setValue("Total new meeting count: " + totalNewMeetings);
  outputRange.offset(rowOffset++, 0).setValue("Total recurring meeting count: " + (totalMeetings - totalNewMeetings));
  outputRange.offset(rowOffset++, 0).setValue("Total demo count: " + statsLedgerG.global.agendaItems.demo);
  outputRange.offset(rowOffset++, 0).setValue("Total pov count: " + statsLedgerG.global.agendaItems.pov);
  outputRange.offset(rowOffset++, 0).setValue("Total workshop count: " + statsLedgerG.global.agendaItems.workshop);
  outputRange.offset(rowOffset++, 0).setValue("Total deep dive count: " + statsLedgerG.global.agendaItems.deepdive);
  

  let ave = Math.round(totalMeetings / userCount);
  let aveN = Math.round(totalNewMeetings / userCount);
  let aveR = Math.round((totalMeetings - totalNewMeetings) / userCount);

  outputRange.offset(rowOffset++, 0).setValue("Average meeting count: " + ave);
  outputRange.offset(rowOffset++, 0).setValue("Average new meeting count: " + aveN);
  outputRange.offset(rowOffset++, 0).setValue("Average recurring meeting count: " + aveR);

  let std = 0;
  let stdN = 0;
  let stdR = 0;
  if (userCount > 1) {

    // Calculate standard deviation
    let squares = 0;
    let squaresN = 0;
    let squaresR = 0;
    for (user in statsLedgerG.users) {

      let t = statsLedgerG.users[user].generalCounts.op + statsLedgerG.users[user].generalCounts.customer +
        statsLedgerG.users[user].generalCounts.partner + statsLedgerG.users[user].generalCounts.lead;
      let tN = statsLedgerG.users[user].generalCounts.op_new + statsLedgerG.users[user].generalCounts.customer_new +
        statsLedgerG.users[user].generalCounts.partner_new + statsLedgerG.users[user].generalCounts.lead_new;

      let tR = t - tN;

      squares = squares + ((t - ave) * (t - ave));
      squaresN = squaresN + ((tN - aveN) * (tN - aveN));
      squaresR = squaresR + ((tR - aveR) * (tR - aveR));

    }
    std = Math.round(Math.sqrt(squares / (userCount - 1)));
    stdN = Math.round(Math.sqrt(squaresN / (userCount - 1)));
    stdR = Math.round(Math.sqrt(squaresR / (userCount - 1)));
    outputRange.offset(rowOffset++, 0).setValue("StdDev meeting count: " + std);
    outputRange.offset(rowOffset++, 0).setValue("StdDev new meeting count: " + stdN);
    outputRange.offset(rowOffset++, 0).setValue("StdDev recurring meeting count: " + stdR);
  }
  rowOffset++;
  rowOffset++;

  outputRange.offset(rowOffset++, 0).setValue("Aggregate Agenda Item Percentages");
  let outputFrameRowOffset = rowOffset;
  chartLedger[chartIndex] = { title: "Aggregate Agenda Item Percentages", rowOffset: rowOffset, rowCount: 0, colOffset: 0, colCount: 2 };
  outputRange.offset(rowOffset, 1).setValue("Percentage of Total Meetings");
  rowOffset++;
  outputRange.offset(rowOffset, 0).setValue("Demo");
  outputRange.offset(rowOffset, 1).setValue((100 * statsLedgerG.global.agendaItems.demo / totalMeetings).toFixed(1));
  rowOffset++;
  outputRange.offset(rowOffset, 0).setValue("Workshop");
  outputRange.offset(rowOffset, 1).setValue((100 * statsLedgerG.global.agendaItems.workshop / totalMeetings).toFixed(1));
  rowOffset++;
  outputRange.offset(rowOffset, 0).setValue("POV");
  outputRange.offset(rowOffset, 1).setValue((100 * statsLedgerG.global.agendaItems.pov / totalMeetings).toFixed(1));
  rowOffset++;
  outputRange.offset(rowOffset, 0).setValue("Deep Dive");
  outputRange.offset(rowOffset, 1).setValue((100 * statsLedgerG.global.agendaItems.deepdive / totalMeetings).toFixed(1));
  rowOffset++;
  chartLedger[chartIndex].rowCount = rowOffset - chartLedger[chartIndex].rowOffset;
  chartIndex++

  rowOffset++;
  if (rowOffset - outputFrameRowOffset < minChartSpacing) {
    rowOffset = outputFrameRowOffset + minChartSpacing;
  }

  outputRange.offset(rowOffset++, 0).setValue("Total External Meeting Counts");
  outputFrameRowOffset = rowOffset;
  chartLedger[chartIndex] = { title: "Total External Meeting Counts", rowOffset: rowOffset, rowCount: 0, colOffset: 0, colCount: 5 };
  outputRange.offset(rowOffset, 1).setValue("Opportunity");
  outputRange.offset(rowOffset, 2).setValue("Customer");
  outputRange.offset(rowOffset, 3).setValue("Partner");
  outputRange.offset(rowOffset, 4).setValue("Lead");
  rowOffset++;
  for (user in statsLedgerG.users) {
    outputRange.offset(rowOffset, 0).setValue(user);
    outputRange.offset(rowOffset, 1).setValue(statsLedgerG.users[user].generalCounts.op);
    outputRange.offset(rowOffset, 2).setValue(statsLedgerG.users[user].generalCounts.customer);
    outputRange.offset(rowOffset, 3).setValue(statsLedgerG.users[user].generalCounts.partner);
    outputRange.offset(rowOffset, 4).setValue(statsLedgerG.users[user].generalCounts.lead);
    rowOffset++;
  }
  chartLedger[chartIndex].rowCount = rowOffset - chartLedger[chartIndex].rowOffset;
  chartIndex++

  rowOffset++;
  if (rowOffset - outputFrameRowOffset < minChartSpacing) {
    rowOffset = outputFrameRowOffset + minChartSpacing;
  }

  outputRange.offset(rowOffset++, 0).setValue("New External Meeting Counts");
  outputFrameRowOffset = rowOffset;
  chartLedger[chartIndex] = { title: "New External Meeting Counts", rowOffset: rowOffset, rowCount: 0, colOffset: 0, colCount: 5 };
  outputRange.offset(rowOffset, 1).setValue("Opportunity");
  outputRange.offset(rowOffset, 2).setValue("Customer");
  outputRange.offset(rowOffset, 3).setValue("Partner");
  outputRange.offset(rowOffset, 4).setValue("Lead");
  rowOffset++;
  for (user in statsLedgerG.users) {
    outputRange.offset(rowOffset, 0).setValue(user);
    outputRange.offset(rowOffset, 1).setValue(statsLedgerG.users[user].generalCounts.op_new);
    outputRange.offset(rowOffset, 2).setValue(statsLedgerG.users[user].generalCounts.customer_new);
    outputRange.offset(rowOffset, 3).setValue(statsLedgerG.users[user].generalCounts.partner_new);
    outputRange.offset(rowOffset, 4).setValue(statsLedgerG.users[user].generalCounts.lead_new);
    rowOffset++;
  }
  chartLedger[chartIndex].rowCount = rowOffset - chartLedger[chartIndex].rowOffset;
  chartIndex++

  rowOffset++;
  if (rowOffset - outputFrameRowOffset < minChartSpacing) {
    rowOffset = outputFrameRowOffset + minChartSpacing;
  }

  outputRange.offset(rowOffset++, 0).setValue("Recurring External Meeting Counts");
  outputFrameRowOffset = rowOffset;
  chartLedger[chartIndex] = { title: "Recurring External Meeting Counts", rowOffset: rowOffset, rowCount: 0, colOffset: 0, colCount: 5 };
  outputRange.offset(rowOffset, 1).setValue("Opportunity");
  outputRange.offset(rowOffset, 2).setValue("Customer");
  outputRange.offset(rowOffset, 3).setValue("Partner");
  outputRange.offset(rowOffset, 4).setValue("Lead");
  rowOffset++;
  for (user in statsLedgerG.users) {
    outputRange.offset(rowOffset, 0).setValue(user);
    outputRange.offset(rowOffset, 1).setValue(statsLedgerG.users[user].generalCounts.op - statsLedgerG.users[user].generalCounts.op_new);
    outputRange.offset(rowOffset, 2).setValue(statsLedgerG.users[user].generalCounts.customer - statsLedgerG.users[user].generalCounts.customer_new);
    outputRange.offset(rowOffset, 3).setValue(statsLedgerG.users[user].generalCounts.partner - statsLedgerG.users[user].generalCounts.partner_new);
    outputRange.offset(rowOffset, 4).setValue(statsLedgerG.users[user].generalCounts.lead - statsLedgerG.users[user].generalCounts.lead_new);
    rowOffset++;
  }
  chartLedger[chartIndex].rowCount = rowOffset - chartLedger[chartIndex].rowOffset;
  chartIndex++


  if (rowOffset - outputFrameRowOffset < minChartSpacing) {
    rowOffset = outputFrameRowOffset + minChartSpacing;
  }

 outputRange.offset(rowOffset++, 0).setValue("Agenda Item Counts");
  outputFrameRowOffset = rowOffset;
  chartLedger[chartIndex] = { title: "Agenda Item Counts", rowOffset: rowOffset, rowCount: 0, colOffset: 0, colCount: 5 };
  outputRange.offset(rowOffset, 1).setValue("Demo");
  outputRange.offset(rowOffset, 2).setValue("Workshop");
  outputRange.offset(rowOffset, 3).setValue("POV");
  outputRange.offset(rowOffset, 4).setValue("Deep Dive");
  rowOffset++;
  for (user in statsLedgerG.users) {
    outputRange.offset(rowOffset, 0).setValue(user);
    outputRange.offset(rowOffset, 1).setValue(statsLedgerG.users[user].generalCounts.agendaItems.demo);
    outputRange.offset(rowOffset, 2).setValue(statsLedgerG.users[user].generalCounts.agendaItems.workshop);
    outputRange.offset(rowOffset, 3).setValue(statsLedgerG.users[user].generalCounts.agendaItems.pov);
    outputRange.offset(rowOffset, 4).setValue(statsLedgerG.users[user].generalCounts.agendaItems.deepdive);
    rowOffset++;
  }
  chartLedger[chartIndex].rowCount = rowOffset - chartLedger[chartIndex].rowOffset;
  chartIndex++

  if (rowOffset - outputFrameRowOffset < minChartSpacing) {
    rowOffset = outputFrameRowOffset + minChartSpacing;
  }

  outputRange.offset(rowOffset++, 0).setValue("------------------------------------------------------------------------");
  outputRange.offset(rowOffset++, 0).setValue("                     Staff Member Meeting Counts");
  outputRange.offset(rowOffset++, 0).setValue("------------------------------------------------------------------------");
  rowOffset++

  for (user in statsLedgerG.users) { // user is email

    outputRange.offset(rowOffset++, 0).setValue(user + " Stats");
    outputRange.offset(rowOffset++, 0).setValue("-----------------------");
    outputRange.offset(rowOffset++, 0).setValue("Engaged Opportunities: " + statistics[user].ops);
    outputRange.offset(rowOffset++, 0).setValue("Engaged Customers (with no op): " + statistics[user].customers);
    outputRange.offset(rowOffset++, 0).setValue("Engaged Partners: " + statistics[user].partners);
    outputRange.offset(rowOffset++, 0).setValue("Demos: " + statsLedgerG.users[user].generalCounts.agendaItems.demo);
    outputRange.offset(rowOffset++, 0).setValue("Workshops: " + statsLedgerG.users[user].generalCounts.agendaItems.workshop);
    outputRange.offset(rowOffset++, 0).setValue("POV Meetings: " + statsLedgerG.users[user].generalCounts.agendaItems.pov);
    outputRange.offset(rowOffset++, 0).setValue("Deep Dives: " + statsLedgerG.users[user].generalCounts.agendaItems.deepdive);
    let message = "";
    if (statistics[user].new_meetings < aveN) {
      if (aveN - statistics[user].new_meetings > stdN) {
        message = "Number of new meetings: " + statistics[user].new_meetings + ", more than a standard deviation BELOW the mean of " + aveN;
      }
      else {
        message = "Number of new meetings: " + statistics[user].new_meetings + ", which is BELOW the mean of " + aveN;
      }
    }
    else {
      if (statistics[user].new_meetings - aveN > stdN) {
        message = "Number of new meetings: " + statistics[user].new_meetings + ", more than a standard deviation above than the mean of " + aveN;
      }
      else {
        message = "Number of new meetings: " + statistics[user].new_meetings + ", which is above the mean of " + aveN;
      }
    }
    outputRange.offset(rowOffset++, 0).setValue(message);
    if (statistics[user].recurring_meetings < aveR) {
      if (aveR - statistics[user].recurring_meetings > stdR) {
        message = "Number of recurring meetings: " + statistics[user].recurring_meetings + ", more than a standard deviation BELOW the mean of " + aveR;
      }
      else {
        message = "Number of recurring meetings: " + statistics[user].recurring_meetings + ", which is BELOW the mean of " + aveR;
      }
    }
    else {
      if (statistics[user].recurring_meetings - aveR > stdR) {
        message = "Number of recurring meetings: " + statistics[user].recurring_meetings + ", more than a standard deviation above than the mean of " + aveR;
      }
      else {
        message = "Number of recurring meetings: " + statistics[user].recurring_meetings + ", which is above the mean of " + aveR;
      }
    }
    outputRange.offset(rowOffset++, 0).setValue(message);
    outputRange.offset(rowOffset++, 0).setValue("Total number of contacts attending new meetings (moving the needle forward): " + statistics[user].attendees_new);
    outputRange.offset(rowOffset++, 0).setValue("Total number of contacts attending recurring meetings (status calls): " + statistics[user].attendees_recurring);

    rowOffset = rowOffset + 2;
    outputRange.offset(rowOffset++, 0).setValue(user + " Event Counts");
    outputFrameRowOffset = rowOffset;
    chartLedger[chartIndex] = { title: user + " Event Counts", rowOffset: rowOffset, rowCount: 0, colOffset: 0, colCount: 6 };
    outputRange.offset(rowOffset, 1).setValue("Opportunity");
    outputRange.offset(rowOffset, 2).setValue("Customer");
    outputRange.offset(rowOffset, 3).setValue("Partner");
    outputRange.offset(rowOffset, 4).setValue("Lead");
    outputRange.offset(rowOffset, 5).setValue("Recurring");
    rowOffset++;
    for (op in statsLedgerG.users[user].ops) {
      outputRange.offset(rowOffset, 0).setValue(op);
      outputRange.offset(rowOffset, 1).setValue(statsLedgerG.users[user].ops[op].new_count);
      outputRange.offset(rowOffset, 2).setValue(0);
      outputRange.offset(rowOffset, 3).setValue(0);
      outputRange.offset(rowOffset, 4).setValue(0);
      outputRange.offset(rowOffset, 5).setValue(statsLedgerG.users[user].ops[op].recurring_count);
      rowOffset++;
    }
    for (a in statsLedgerG.users[user].customers) {
      outputRange.offset(rowOffset, 0).setValue(a);
      outputRange.offset(rowOffset, 1).setValue(0);
      outputRange.offset(rowOffset, 2).setValue(statsLedgerG.users[user].customers[a].new_count);
      outputRange.offset(rowOffset, 3).setValue(0);
      outputRange.offset(rowOffset, 4).setValue(0);
      outputRange.offset(rowOffset, 5).setValue(statsLedgerG.users[user].customers[a].recurring_count);
      rowOffset++;
    }
    for (a in statsLedgerG.users[user].partners) {
      outputRange.offset(rowOffset, 0).setValue(a);
      outputRange.offset(rowOffset, 1).setValue(0);
      outputRange.offset(rowOffset, 2).setValue(0);
      outputRange.offset(rowOffset, 3).setValue(statsLedgerG.users[user].partners[a].new_count);
      outputRange.offset(rowOffset, 4).setValue(0);
      outputRange.offset(rowOffset, 5).setValue(statsLedgerG.users[user].partners[a].recurring_count);
      rowOffset++;
    }
    for (l in statsLedgerG.users[user].leads) {
      outputRange.offset(rowOffset, 0).setValue(l);
      outputRange.offset(rowOffset, 1).setValue(0);
      outputRange.offset(rowOffset, 2).setValue(0);
      outputRange.offset(rowOffset, 3).setValue(0);
      outputRange.offset(rowOffset, 4).setValue(statsLedgerG.users[user].leads[l].new_count);
      outputRange.offset(rowOffset, 5).setValue(statsLedgerG.users[user].leads[l].recurring_count);
      rowOffset++;
    }
    chartLedger[chartIndex].rowCount = rowOffset - chartLedger[chartIndex].rowOffset;
    chartIndex++

    rowOffset++;
    if (rowOffset - outputFrameRowOffset < minChartSpacing) {
      rowOffset = outputFrameRowOffset + minChartSpacing;
    }

    rowOffset = rowOffset + 2;
    outputRange.offset(rowOffset++, 0).setValue(user + " Agenda Item Counts");
    outputFrameRowOffset = rowOffset;
    chartLedger[chartIndex] = { title: user + " Agenda Item Counts", rowOffset: rowOffset, rowCount: 0, colOffset: 0, colCount: 5 };
    outputRange.offset(rowOffset, 1).setValue("Demos");
    outputRange.offset(rowOffset, 2).setValue("Workshops");
    outputRange.offset(rowOffset, 3).setValue("POVs");
    outputRange.offset(rowOffset, 4).setValue("Deep Dives");
    rowOffset++;
    for (op in statsLedgerG.users[user].ops) {
      outputRange.offset(rowOffset, 0).setValue(op);
      outputRange.offset(rowOffset, 1).setValue(statsLedgerG.users[user].ops[op].agendaItems.demo);
      outputRange.offset(rowOffset, 2).setValue(statsLedgerG.users[user].ops[op].agendaItems.workshop);
      outputRange.offset(rowOffset, 3).setValue(statsLedgerG.users[user].ops[op].agendaItems.pov);
      outputRange.offset(rowOffset, 4).setValue(statsLedgerG.users[user].ops[op].agendaItems.deepdive);
      rowOffset++;
    }
    for (a in statsLedgerG.users[user].customers) {
      outputRange.offset(rowOffset, 0).setValue(a);
      outputRange.offset(rowOffset, 1).setValue(statsLedgerG.users[user].customers[a].agendaItems.demo);
      outputRange.offset(rowOffset, 2).setValue(statsLedgerG.users[user].customers[a].agendaItems.workshop);
      outputRange.offset(rowOffset, 3).setValue(statsLedgerG.users[user].customers[a].agendaItems.pov);
      outputRange.offset(rowOffset, 4).setValue(statsLedgerG.users[user].customers[a].agendaItems.deepdive);
      rowOffset++;
    }
    for (a in statsLedgerG.users[user].partners) {
      outputRange.offset(rowOffset, 0).setValue(a);
      outputRange.offset(rowOffset, 1).setValue(statsLedgerG.users[user].partners[a].agendaItems.demo);
      outputRange.offset(rowOffset, 2).setValue(statsLedgerG.users[user].partners[a].agendaItems.workshop);
      outputRange.offset(rowOffset, 3).setValue(statsLedgerG.users[user].partners[a].agendaItems.pov);
      outputRange.offset(rowOffset, 4).setValue(statsLedgerG.users[user].partners[a].agendaItems.deepdive);
      rowOffset++;
    }
    chartLedger[chartIndex].rowCount = rowOffset - chartLedger[chartIndex].rowOffset;
    chartIndex++

    rowOffset++;
    if (rowOffset - outputFrameRowOffset < minChartSpacing) {
      rowOffset = outputFrameRowOffset + minChartSpacing;
    }
  }

  outputRange.offset(rowOffset++, 0).setValue("------------------------------------------------------------------------");
  outputRange.offset(rowOffset++, 0).setValue("                             Contact Tracing");
  outputRange.offset(rowOffset++, 0).setValue("------------------------------------------------------------------------");
  rowOffset++



  for (user in statsLedgerG.users) {

    outputRange.offset(rowOffset++, 0).setValue("-------------------  for " + user + "  -------------------");
    rowOffset = rowOffset + 2;

    outputRange.offset(rowOffset++, 1).setValue("-------------------  Opportunity Contact Counts  -------------------");
    rowOffset = rowOffset + 2;

    let maxFrameRows = 0;
    let outputFrameColOffset = 1;
    outputFrameRowOffset = rowOffset;
    for (op in statsLedgerG.users[user].ops) {
      printContactHeader_(outputRange, rowOffset, outputFrameColOffset, user, op);
      rowOffset = rowOffset + 2;
      let contactArray = [];
      for (a in statsLedgerG.users[user].ops[op].attendees) {
        contactArray.push({ n: a, u: statsLedgerG.users[user].ops[op].attendees[a].new_count, 
        r: statsLedgerG.users[user].ops[op].attendees[a].recurring_count, ai: statsLedgerG.users[user].ops[op].attendees[a].agendaItems })
      }
      contactArray.sort(compareContacts_);
      for (let i = 0; i < contactArray.length; i++) {
        printContactStats_(outputRange, rowOffset, outputFrameColOffset, contactArray[i]);
        rowOffset++;
      }
      if (rowOffset - outputFrameRowOffset > maxFrameRows) {
        maxFrameRows = rowOffset - outputFrameRowOffset;
      }
      rowOffset = outputFrameRowOffset
      outputFrameColOffset = outputFrameColOffset + 9;
    }

    rowOffset = outputFrameRowOffset + maxFrameRows + 1;
    outputFrameColOffset = 1;
    maxFrameRows = 0;
    outputRange.offset(rowOffset++, 1).setValue("------------------- Account Contact Counts (No Opportunity registered) -------------------");
    rowOffset = rowOffset + 2;
    outputFrameRowOffset = rowOffset; // In case nothing is in the next loop
    for (account in statsLedgerG.users[user].customers) {
      printContactHeader_(outputRange, rowOffset, outputFrameColOffset, user, account);
      rowOffset = rowOffset + 2;
      let contactArray = [];
      for (a in statsLedgerG.users[user].customers[account].attendees) {
        contactArray.push({ n: a, u: statsLedgerG.users[user].customers[account].attendees[a].new_count, 
        r: statsLedgerG.users[user].customers[account].attendees[a].recurring_count, ai: statsLedgerG.users[user].customers[account].attendees[a].agendaItems })
      }
      contactArray.sort(compareContacts_);
      for (let i = 0; i < contactArray.length; i++) {
        printContactStats_(outputRange, rowOffset, outputFrameColOffset, contactArray[i]);
        rowOffset++;
      }
      if (rowOffset - outputFrameRowOffset > maxFrameRows) {
        maxFrameRows = rowOffset - outputFrameRowOffset;
      }
      rowOffset = outputFrameRowOffset
      outputFrameColOffset = outputFrameColOffset + 9;
    }

    rowOffset = outputFrameRowOffset + maxFrameRows + 1;
    outputFrameColOffset = 1;
    maxFrameRows = 0;
    outputRange.offset(rowOffset++, 1).setValue("-------------------  Partner Contact Counts -------------------");
    rowOffset = rowOffset + 2;
    outputFrameRowOffset = rowOffset;
    for (account in statsLedgerG.users[user].partners) {
      printContactHeader_(outputRange, rowOffset, outputFrameColOffset, user, account);
      rowOffset = rowOffset + 2;
      let contactArray = [];
      for (a in statsLedgerG.users[user].partners[account].attendees) {
        contactArray.push({ n: a, u: statsLedgerG.users[user].partners[account].attendees[a].new_count, 
        r: statsLedgerG.users[user].partners[account].attendees[a].recurring_count, ai: statsLedgerG.users[user].partners[account].attendees[a].agendaItems })
      }
      contactArray.sort(compareContacts_);
      for (let i = 0; i < contactArray.length; i++) {
        printContactStats_(outputRange, rowOffset, outputFrameColOffset, contactArray[i]);
        rowOffset++;
      }

      if (rowOffset - outputFrameRowOffset > maxFrameRows) {
        maxFrameRows = rowOffset - outputFrameRowOffset;
      }
      rowOffset = outputFrameRowOffset
      outputFrameColOffset = outputFrameColOffset + 9;
    }

    rowOffset = outputFrameRowOffset + maxFrameRows + 1;

  }


  //
  // Build Charts
  //
  logOneCol("Generating charts in the " + MEETING + " tab of the stats output sheet");
  for (let ci = 0; ci < chartIndex; ci++) {

    //Logger.log(JSON.stringify(chartLedger[ci]));

    let dataRange = sheet.getRange(MEETING_COUNTS_ORIGIN_ROW + chartLedger[ci].rowOffset, MEETING_COUNTS_ORIGIN_COL + chartLedger[ci].colOffset,
      chartLedger[ci].rowCount, chartLedger[ci].colCount);
    makeStackedColumnChart(chartLedger[ci].title, sheet, dataRange, MEETING_COUNTS_ORIGIN_ROW + chartLedger[ci].rowOffset,
      MEETING_COUNTS_ORIGIN_COL + chartLedger[ci].colOffset + chartLedger[ci].colCount); // spread out overlapping charts over 4 columns
  }
}

function compareContacts_(a, b) {
  // u is "new meeting count" for that contact
  if (a.u > b.u) return -1;
  if (a.u < b.u) return 1;
  return 0;
}

function deleteCharts_(sheet) {
  //sheet = SpreadsheetApp.getActive().getSheetByName(STATS);
  var charts = sheet.getCharts();
  for (var i in charts) {
    sheet.removeChart(charts[i]);
  }
}

function incrementStat_(obj, key, eventInfo) {
  if (obj[key]) {
    if (eventInfo[EVENT_RECURRING]) {
      obj[key].recurring_count++;
    }
    else {
      obj[key].new_count++;
    }
  }
  else {
    let agendaItems = {demo : 0, pov : 0, workshop : 0, deepdive : 0 };
    if (eventInfo[EVENT_RECURRING]) {
      obj[key] = { new_count: 0, recurring_count: 1, agendaItems: agendaItems};
    }
    else {
      obj[key] = { new_count: 1, recurring_count: 0, agendaItems: agendaItems};
    }
  }
  incrementAgendaItemCounts_(obj[key].agendaItems, eventInfo);
}

function printContactHeader_(outputRange, rowOffset, outputFrameColOffset, user, relatedTo) {
  outputRange.offset(rowOffset++, outputFrameColOffset).setValue(relatedTo + " (" + user + ")");
  outputRange.offset(rowOffset, outputFrameColOffset + 1).setValue("New Meeting");
  outputRange.offset(rowOffset, outputFrameColOffset + 2).setValue("Recurring");
  outputRange.offset(rowOffset, outputFrameColOffset + 3).setValue("Demo");
  outputRange.offset(rowOffset, outputFrameColOffset + 4).setValue("POV");
  outputRange.offset(rowOffset, outputFrameColOffset + 5).setValue("Workshop");
  outputRange.offset(rowOffset, outputFrameColOffset + 6).setValue("Deepdive");
}

function printContactStats_(outputRange, rowOffset, outputFrameColOffset, contact) {
  outputRange.offset(rowOffset, outputFrameColOffset).setValue(contact.n);
  outputRange.offset(rowOffset, outputFrameColOffset + 1).setValue(contact.u);
  outputRange.offset(rowOffset, outputFrameColOffset + 2).setValue(contact.r);
  outputRange.offset(rowOffset, outputFrameColOffset + 3).setValue(contact.ai.demo);
  outputRange.offset(rowOffset, outputFrameColOffset + 4).setValue(contact.ai.pov);
  outputRange.offset(rowOffset, outputFrameColOffset + 5).setValue(contact.ai.workshop);
  outputRange.offset(rowOffset, outputFrameColOffset + 6).setValue(contact.ai.deepdive);
}


function logStamp(title) {

  let logSheet = SpreadsheetApp.getActive().getSheetByName(LOG_TAB);
  var logLastRow = logSheet.getLastRow();
  AM_LOG = logSheet.getRange(logLastRow + 2, 1); // Leave an empty row divider
  AM_LOG_ROW = 0;

  AM_LOG.offset(AM_LOG_ROW, 0).setValue(title + " " + new Date().toLocaleDateString() + " " + new Date().toLocaleTimeString());
  AM_LOG.offset(AM_LOG_ROW + 1, 0).setValue("---------------------------------------------------------------");
  AM_LOG_ROW += 2;
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