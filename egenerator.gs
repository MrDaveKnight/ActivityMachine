function build_se_events() {
  
  // Objective
  //
  // Automatically create SE Activities (Field Events in Salesforce) from calendar invites.
  //
  // Design Strategy
  //
  // Parse each calendar invite looking for:
  // - The organizer
  //.- The person invited by the organizer
  // - All the attendee emails
  // - The products beginning discussed
  // - The "general" type of activity (workshop, demo, presentation, more on this later)
  //
  // Determine the CUSTOMERS associated with all of the attendees. Determine which, if any,
  // customer has the majority of attendees. Assume this is the customer the invite is "targeted" for. Others
  // are partners or contractors. If no clear majority, assume the invite is a type of workshop
  // or other multi-customer event. 
  //
  // Determine what PRODUCTS are being discussed. Look first for any tags of the
  // form 'Product:<product>'(spaces around colon are optional).
  // If no tags, look for <product> tokens in the subject. If none, look for <product> tokens in the 
  // description. The result will be one or more products. 
  //
  // For a targeted customer invite, search for OPPORTUNITIES that match the targeted customer and products. For each
  // customer/product opportunity, create an SE activity for that opportunity. If no opportunities are located, 
  // create an SE activity for that targeted customer.   
  //
  // For a workshop invite, look for opportunities for each customer/product pair. When found, create an SE activity
  // for that opportunity. If an opportunity is not found for any customer, create an SE activity for that customer. 
  //
  // When searching for opportunities, if two or more are found that match the customer and product, 
  // select the one with close date in closest proximity to the invite date.
  // 
  // When creating an SE activity for an opportunity, determine in what phase of the opporunity the calendar invite
  // started: Discovery, Alignment (Technical & Business Validation till pre-close), Closed/Won or 
  // Closed/Lost (all the Closed lost stages). 
  // 
  // Set the Field Event's 'Opportunity Stage' as follows:
  // - If in Discovery: "Discovery & Qualification"
  // - If in Alignment: "Technical & Business Validation"
  // - If in Closed/Won: "Closed Won"
  // - If in Closed/Lost: "Closed Lost"
  //
  // Discovery & Qualification is the default. Only two dates are needed, when the op moved to 
  // Technical & Business Validation, and when it closed.
  //
  // Set the Field Event's 'Meeting Type" based on the Opportunity Stage, searching first for specific tags of
  // the form "Activity:<activity>", e.g. Activity:demo (spaces around colon are optional). Some meeting 
  // types will force a reset of the stage (due to Salesforce restrictions on Meeting Types for each Stage)
  //
  // If no tags, search for key words in a specific sequence to derive an activity.
  // 
  // For reasons I won't go into here, only a subset of the Meeting Types in the Saleforce SE Activity record are
  // eligable for selection. In addition, some are "re-mapped" a bit to have a slightly different meaning.
  //
  // Here are the "Meeting Types" this system will identify. They are ordered generally based on the prep, effort and skill 
  // required of an SE to do it:
  // 
  //    - Shadow
  //    - Social (pre or post-sales activity)
  //    - Discovery (just business or technical discovery or alignment, adhoc Q&A. If a presentation is given, use Presentation)
  //    - Presentation (corp pitch, product overview, practitioner briefing, etc., will likely include discovery)
  //    - Roadmap (special presentation of what's coming)
  //    - Training (post-sales activity)
  //    - HealthCheck
  //    - Support (pre or post-sales activity)
  //    - Demo (will likely include discovery and presentation)
  //    - Workshop (will likely include demo and presentation)
  //    - DeepDive (will likely include demo and presentation)
  //    - POV
  //    - Pilot   
  // 
  //
  //  Here are the mappings from internal to Salesforce Meeting Types. Note that some Meeting Types are only eligible 
  //  in certain Opportunity Stages, so they will invoke a "forced stage SWITCH"
  //
  //  Discovery & Qualification
  //    - Shadow -> Shadow
  //    - Social -> Happy Hour (any social event)
  //    - Discovery -> Discovery (you are just there to ask and answer questions)
  //    - Presentation -> Product Overview (may include Corporate Pitch, Practioner Briefing, EBC stuff) 
  //    - Roadmap -> Product Roadmap (forced SWITCH of Stage to Tecnical & Business Validation)
  //    - Training -> Technical Office Hours (assumed to be training on OSS for reciprocity)
  //    - HealthCheck -> Health Check (forced SWITCH of Stage to Tecnical & Business Validation)
  //    - Support -> Support (assumed to be pre-sales technical support of OSS for reciprocity)
  //    - Demo -> Demo
  //    - Workshop -> Standard Workshop (forced SWITCH of Stage to Tecnical & Business Validation)
  //    - DeepDive -> Product Deep Dive (forced SWITCH of Stage to Tecnical & Business Validation)
  //    - POV -> Controlled POV (forced SWITCH of Stage to Tecnical & Business Validation)
  //    - Pilot -> Pilot (forced SWITCH of Stage to Success Planning)
  //
  //  Technical & Business Validation (which is all non-closed stages unless there is a pilot)
  //    - Shadow -> Shadow
  //    - Social -> Happy Hour (any social event)
  //    - Discovery -> Discovery (forced SWITCH of Stage to Discovery & Qualification)
  //    - Presentation -> Product Overview (may include Corporate Pitch, Practioner Briefing, EBC stuff) 
  //    - RoadMap -> Product Roadmap
  //    - Training -> Virtual SE Office Hours (didn't want to pick Early Education EBC)
  //    - HealthCheck -> Health Check 
  //    - Support -> Support (assumed to be pre-sales technical support of OSS for reciprocity)
  //    - Demo -> Demo
  //    - Workshop -> Standard Workshop (forced SWITCH of Stage to Tecnical & Business Validation)
  //    - DeepDive -> Product Deep Dive 
  //    - POV -> Controlled POV
  //    - Pilot -> Pilot (forced SWITCH of Stage to Success Planning)
  //
  //  Closed/Won
  //    - Shadow -> Shadow
  //    - Social -> Customer Business Review
  //    - Discovery -> Customer Business Review
  //    - Presentation -> Customer Business Review
  //    - RoadMap -> Customer Business Review
  //    - Training -> Training
  //    - HealthCheck -> Customer Business Review
  //    - Support -> Implementation
  //    - Demo -> Customer Business Review
  //    - Workshop -> Training
  //    - DeepDive -> Training
  //    - POV -> Training
  //    - Pilot -> Training
  //
  //  Closed/Lost
  //    There are no allowed meeting types. Post any activity against the ACCOUNT with Discovery & Qualification stage
  

  
  // Information for accounts, staff, opportunities and calendar invites
  // is loaded from tabs in the spreadsheet into two-dimentional arrays with
  // this naming convention -
  //
  //      <info-type>Info 
  //
  
  clearTab_(EVENTS, EVENT_HEADER);
  
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
  // Load Staff Info - SEs and Reps
  //
  
  var sheet = SpreadsheetApp.getActive().getSheetByName(STAFF);
  var rangeData = sheet.getDataRange();
  var slc = rangeData.getLastColumn();
  var slr = rangeData.getLastRow();
  var scanRange = sheet.getRange(2,1, slr-1, slc);
  var staffInfo = scanRange.getValues();
  
  if (slc < STAFF_COLUMNS) {
    Logger.log("ERROR: Imported Staff info was only " + slc + " fields wide. Not enough! Something's not good.");
    return;
  }
  
  // Build the email to user id mapping  
  for (var j = 0 ; j < slr - 1; j++) {
    if (!staffInfo[j][STAFF_EMAIL]) {
      Logger.log("WARNING:" + staffInfo[j][STAFF_NAME] + " has no email!");
      continue;
    }
    staffNameToEmailMap[staffInfo[j][STAFF_NAME].trim()] = staffInfo[j][STAFF_EMAIL].trim();
    staffEmailToIdMap[staffInfo[j][STAFF_EMAIL].trim()] = staffInfo[j][STAFF_ID].trim();
    staffEmailToRoleMap[staffInfo[j][STAFF_EMAIL].trim()] = staffInfo[j][STAFF_ROLE].trim();
    //Logger.log("DEBUG: " + staffInfo[j][STAFF_EMAIL] + " -> " + staffInfo[j][STAFF_ID]);
    
  }
  
  //
  // Load Opportunity Info
  //
  
  sheet = SpreadsheetApp.getActive().getSheetByName(OPPORTUNITIES);
  rangeData = sheet.getDataRange();
  var olc = rangeData.getLastColumn();
  var olr = rangeData.getLastRow();
  scanRange = sheet.getRange(2,1, olr-1, olc);
  var opInfo = scanRange.getValues();
  
  if (olc < OP_COLUMNS) {
    Logger.log("ERROR: Imported opportunity info was only " + olc + " fields wide. Not enough! Something needs to be fixed.");
    return;
  }
  
  // Build the opportunity indexes  
  //
  // There may be many opportunities for a particular customer and product (renewal, new business, services, etc);
  // We want to give priority to the op type, and the date. Pick the earlier op for the meeting. Pick an op that brings in 
  // new business over services and renewals
  // Also track Ops that are closed/lost. We don't want those to be the default even if they were first.
  let opTypeIndexedByCustomerAndProduct = {}; // For op selection priority DEPRECATED
  let opStageIndexedByCustomerAndProduct = {}; // For op selection priority 
  let primaryOpTypeIndexedByCustomer = {}; // For selecting the primary opportunity out of a set with different products
  let primaryOpStageIndexedByCustomer = {}; // Don't want Closed/Lost to be a default
  for (j = 0 ; j < olr - 1; j++) {
    let scanResults = lookForProducts_(opInfo[j][OP_NAME]);
    let productKey = makeProductKey_(scanResults, opInfo[j][OP_PRIMARY_PRODUCT]);
    if (productKey == "-") {
      Logger.log("WARNING: " + opInfo[j][OP_NAME] + " has no product!");
    }
    
 
    
    // Important Note:
    // The OP_ACCOUNT_ID provided in the opportunity record does NOT have the
    // extra 3 characters at the end needed to make it a so called "18 Digit Account ID".
    // However, the account records we use to find accounts from emails DOES
    // use the full "18 Digit Account ID". Since these account IDs need to sync
    // up at some point, we will be stripping off the 3 character postfix when
    // we create the keys from emails. We don't do that here mind you (they are already gone),
    // but later we will. Just a heads up.
    let key = opInfo[j][OP_ACCOUNT_ID] + productKey;
    
    //REMOVEME
    if ("Jet.com Nomad Enterprise Renewal 2020" == opInfo[j][OP_NAME]) {
      Logger.log("DEBUG: Jet:" + key);
    }
    
    if (!numberOfOpsByCustomer[opInfo[j][OP_ACCOUNT_ID]]) {
      numberOfOpsByCustomer[opInfo[j][OP_ACCOUNT_ID]] = 1;
    }
    else {
      numberOfOpsByCustomer[opInfo[j][OP_ACCOUNT_ID]]++;
    }
    
    if (opByCustomerAndProduct[key]) {
      // Use earliest opportunity. They come in sorted on close date. 
      // However, factor in the type of op.
      
      // Tried "Favor new business over renewal", but that doesn't seem to make sense.
      // FIXME - we need to add "renewal" keyword to invites that are for renewal. Important metric
      // We could also try to figure out contacts associated with an opportunity? Use the invitees as a clue to the op?
      
      /*
      if ((opInfo[j][OP_TYPE] == "New Business" && opTypeIndexedByCustomerAndProduct[key] != "New Business") ||
      (opInfo[j][OP_TYPE] == "Expansion" && opTypeIndexedByCustomerAndProduct[key] != "New Business" && 
      opTypeIndexedByCustomerAndProduct[key] != "Expansion") ||
      (opInfo[j][OP_TYPE] == "Add-on" && opTypeIndexedByCustomerAndProduct[key] != "New Business" && 
      opTypeIndexedByCustomerAndProduct[key] != "Expansion" && opTypeIndexedByCustomerAndProduct[key] != "Add-on") ||
      (opInfo[j][OP_TYPE] == "Renewal" && opTypeIndexedByCustomerAndProduct[key] == "Services")) {  
      
      if (IS_TRACE_ACCOUNT_ON && TRACE_ACCOUNT_ID == opInfo[j][OP_ACCOUNT_ID]) {
      Logger.log("TRACE Account: Tracking op " + opInfo[j][OP_ID] + ":" + opInfo[j][OP_TYPE] + productKey);
      }
      
      opByCustomerAndProduct[key] = opInfo[j][OP_ID];
      opTypeIndexedByCustomerAndProduct[key] = opInfo[j][OP_TYPE];
      }
      */
      
      // Make default an active op if available
      if (opInfo[j][OP_STAGE].indexOf("Closed") == -1 && opStageIndexedByCustomerAndProduct[key].indexOf("Closed") != -1) {
        opByCustomerAndProduct[key] = opInfo[j][OP_ID];
        opTypeIndexedByCustomerAndProduct[key] = opInfo[j][OP_TYPE];
        opStageIndexedByCustomerAndProduct[key] = opInfo[j][OP_STAGE];
      }
    }
    else {
    
      
      if (IS_TRACE_ACCOUNT_ON && TRACE_ACCOUNT_ID == opInfo[j][OP_ACCOUNT_ID]) {
        Logger.log("TRACE Account: Tracking op " + opInfo[j][OP_ID] + ":" + opInfo[j][OP_TYPE] + productKey);
      }
      
      opByCustomerAndProduct[key] = opInfo[j][OP_ID];
      opTypeIndexedByCustomerAndProduct[key] = opInfo[j][OP_TYPE];
      opStageIndexedByCustomerAndProduct[key] = opInfo[j][OP_STAGE];
    }
    if (!primaryOpByCustomer[opInfo[j][OP_ACCOUNT_ID]]) {
      if (IS_TRACE_ACCOUNT_ON && TRACE_ACCOUNT_ID == opInfo[j][OP_ACCOUNT_ID]) {
        Logger.log("TRACE Account: Tracking primary op "+ opInfo[j][OP_ID] + ":" + opInfo[j][OP_TYPE] + productKey);
      }
      primaryOpByCustomer[opInfo[j][OP_ACCOUNT_ID]] = opInfo[j][OP_ID]; // Pick an op for all invites with no product specified to go to
      primaryOpTypeIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] = opInfo[j][OP_TYPE];
      primaryOpStageIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] = opInfo[j][OP_STAGE];
      
    }
    else if (opInfo[j][OP_STAGE].indexOf("Closed") == -1 &&  primaryOpStageIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]].indexOf("Closed") != -1) {
      primaryOpByCustomer[opInfo[j][OP_ACCOUNT_ID]] = opInfo[j][OP_ID]; // Pick an op for all invites with no product specified to go to
      primaryOpTypeIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] = opInfo[j][OP_TYPE];
      primaryOpStageIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] = opInfo[j][OP_STAGE];
    }
    
    
    /*
    else if ((primaryOpStageIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] == "Closed/Lost") ||
             (primaryOpTypeIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] == "Services" && opInfo[j][OP_TYPE] != "Services") ||
      (primaryOpTypeIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] == "Renewal" && (opInfo[j][OP_TYPE] != "Services" && opInfo[j][OP_TYPE] != "Services"))) {
        // Make the primary something other than services or renewals if something else exists
        if (IS_TRACE_ACCOUNT_ON && TRACE_ACCOUNT_ID == opInfo[j][OP_ACCOUNT_ID]) {
          Logger.log("TRACE Account: Resetting primary op to "+ opInfo[j][OP_ID] + ":" + opInfo[j][OP_TYPE] + productKey);
        }
        primaryOpByCustomer[opInfo[j][OP_ACCOUNT_ID]] = opInfo[j][OP_ID]; // Pick an op for all invites with no product specified to go to   
        primaryOpTypeIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] = opInfo[j][OP_TYPE];
        primaryOpStageIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] = opInfo[j][OP_STAGE];
      } */
      
    productByOp[opInfo[j][OP_ID]] = opInfo[j][OP_PRIMARY_PRODUCT];
    if (opInfo[j][OP_PRIMARY_PRODUCT] != "Terraform" && 
        opInfo[j][OP_PRIMARY_PRODUCT] != "Vault" &&
        opInfo[j][OP_PRIMARY_PRODUCT] != "Consul" &&
        opInfo[j][OP_PRIMARY_PRODUCT] != "Nomad") {
      let prod = opInfo[j][OP_PRIMARY_PRODUCT].toLowerCase();
      if (prod.indexOf("terraform") != -1 || prod.indexOf("tfe") || prod.indexOf("tfc")) {
        productByOp[opInfo[j][OP_ID]] = "Terraform";
      }
      else if (prod.indexOf("v") != -1) {
        productByOp[opInfo[j][OP_ID]] = "Vault";
      }
      else if (prod.indexOf("c") != -1) {
        productByOp[opInfo[j][OP_ID]] = "Consul";
      }
      else if (prod.indexOf("n") != -1) {
        productByOp[opInfo[j][OP_ID]] = "Nomad";
      }
      else {
        productByOp[opInfo[j][OP_ID]] = "N/A";
        Logger.log("WARNING: Primary product for Op " + opInfo[j][OP_NAME] + " wasn't valid: " + opInfo[j][OP_PRIMARY_PRODUCT]);
      }
    }
    let team = {se_primary : opInfo[j][OP_SE_PRIMARY] , se_secondary : opInfo[j][OP_SE_SECONDARY] , rep : opInfo[j][OP_OWNER] };
    accountTeamByOp[opInfo[j][OP_ID]] = team;
  }
  
  //
  // Load Opportunity Stage History
  //
  
  var sheet = SpreadsheetApp.getActive().getSheetByName(HISTORY);
  rangeData = sheet.getDataRange();
  var hlc = rangeData.getLastColumn();
  var hlr = rangeData.getLastRow();
  var scanRange = sheet.getRange(2,1, hlr-1, hlc);
  var historyInfo = scanRange.getValues();
  
  if (hlc < HISTORY_COLUMNS) {
    Logger.log("ERROR: Imported opportunity history was only " + hlc + " fields wide. Not enough! Qu'est-ce que c'est?");
    return;
  }
  
  for (j = 0; j < historyInfo.length; j++) {
    if (accountTeamByOp[historyInfo[j][HISTORY_OP_ID]]) {
      // This op is being tracked
      
      if (!stageMilestonesByOp[historyInfo[j][HISTORY_OP_ID]]) {
        stageMilestonesByOp[historyInfo[j][HISTORY_OP_ID]] = { discovery_ended : new Date(2050, 11, 21), closed_at : new Date(2050, 11, 25), was_won : false };
      }
      
      var milestone = stageMilestonesByOp[historyInfo[j][HISTORY_OP_ID]];
      let stage = historyInfo[j][HISTORY_STAGE_VALUE];
      let stageDate = Date.parse(historyInfo[j][HISTORY_STAGE_DATE]);
      if (stage == "Closed/Won" ||
          stage == "Closed/Lost" ||
          stage == "Closed Lost/Churn" ||
          stage == "Debooking") {
        if (stageDate < milestone.closed_at) {
          milestone.closed_at = stageDate;
        }
        if (stage == "Closed/Won") {
          milestone.was_won = true;
        }
      }
      else if (stage != "Discovery & Qualification" && stage != "None") {
        
        // Stage is "Technical & Business Validation" || "Success Planning" || "Deal Review & Proposal" ||
        // "Negotiation & Legal" || "Bookings Review"
        if (stageDate < milestone.discovery_ended) {
          milestone.discovery_ended = stageDate;
        }
      }
    }
  }
  
  // 
  // Process Calendar invites
  //
  
  // The raw calendar invites are in the Calendar tab.
  sheet = SpreadsheetApp.getActive().getSheetByName('Calendar')
  rangeData = sheet.getDataRange();
  var lastColumn = rangeData.getLastColumn();
  var lastRow = rangeData.getLastRow();
  
  if (lastRow == 1) return; // Empty. Only header
  
  if (lastColumn < CALENDAR_COLUMNS){
    Logger.log("ERROR: Imported Calendar was only " + lastColumn + " fields wide. Not enought! Something bad happened.");
    return;
  }
  scanRange = sheet.getRange(2,1, lastRow-1, lastColumn);
  
  // Suck all the calendar data up into memory.
  var inviteInfo = scanRange.getValues();
  
  // Set event output cursor
  var sheet = SpreadsheetApp.getActive().getSheetByName("EVENTS");
  var elr = sheet.getLastRow();
  var outputRange = sheet.getRange(elr+1,1);
  var outputCursor = {range : outputRange, rowOffset : 0};
  
  // 
  // Do what we came here to do!
  // Convert calendar invites (Calendar tab) to SE events (Events tab)
  //
  
  for (j = 0 ; j < lastRow - 1; j++) {
    
    if (!inviteInfo[j][ASSIGNED_TO] || !inviteInfo[j][ATTENDEE_STR] || !inviteInfo[j][START]) {
      continue;
    }
    
    var attendees = inviteInfo[j][ATTENDEE_STR].split(","); // convert comma separated list of emails to an array
    if (attendees.length == 0) {
      continue; 
    }
    
    if (inviteInfo[j][IS_RECURRING] == "TRUE" && inviteInfo[j][ASSIGNEE_STATUS] == "INVITED" || inviteInfo[j][ASSIGNEE_STATUS] == "NO") {
      // If assigned-to has not actually accepted the meeting, don't use it as an SE Activity
      // Unfortunately, if the assigned-to is also the owner of the recurring event, there is no way (I can find) to determine if 
      // they are actually going or not (unless they delete the invite of course). The getMyStatus() API always returns "OWNER" for the owner,
      // even if they are not going (and the UI indicates it; again, can't find a way to get that info out of the google calendar API.)
      continue;
    }
    
    //
    // Look for "special" known meetings we want to track: Virtual Office Hours
    // 
    
    if (inviteInfo[j][SUBJECT].toLowerCase().indexOf("virtual office hours") != -1) {
      if (inviteInfo[j][ASSIGNEE_STATUS] != "NO") {
        let pi = lookForProducts_(inviteInfo[j][SUBJECT] + " " + inviteInfo[j][DESCRIPTION]);
        createSpecialEvents_(outputCursor, attendees, inviteInfo[j], pi, "Virtual SE Office Hours");
      }
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
    
    var assignedTo = staffEmailToIdMap[inviteInfo[j][ASSIGNED_TO]];
    var productInventory = lookForProducts_(inviteInfo[j][SUBJECT] + " " + inviteInfo[j][DESCRIPTION]);
    
    //
    // Manual Overrides
    //
     
    let isOverrideActive = false;
    
    let parms = SpreadsheetApp.getActive().getSheetByName(RUN_PARMS);
    let overrideRange = parms.getRange(4,8,16,3); // Hardcoded to format in RUN_PARMS
    let overrides = overrideRange.getValues();
    
    for (var row in overrides) {
      if (overrides[row][0]) {      
        
        let account = overrides[row][0];
        let subjectTest = false;
        let emailTest = false;
        
        if (!overrides[row][1] && !overrides[row][2]) {
          continue;
        }
        
        if (overrides[row][1]) {
          let subjectRegex = new RegExp(overrides[row][1]);
          subjectTest = subjectRegex.test(inviteInfo[j][SUBJECT])
        }
        else {
          subjectTest = true; 
        }
        
        if (overrides[row][2]) {
          let emailRegex = new RegExp(overrides[row][2]);
          emailTest = emailRegex.test(inviteInfo[j][ATTENDEE_STR])
        }
        else {
          emailTest = true; 
        }
        
        if (subjectTest && emailTest) {
          
          isOverrideActive = true;
          
          if (overrides[row][3] == "Yes") {
            customer = {};
            customer[account] = 1;
            createAccountEvents_(outputCursor, attendees, customer, inviteInfo[j], productInventory);      
            break;
          }
          
          // Find an opportunity
          var key = account + makeProductKey_(productInventory, 0);
          
          var opId = 0;
          if (opByCustomerAndProduct[key]) {
            opId = opByCustomerAndProduct[key];
          }
          
          if (opId) {
            let milestones = { discovery_ended : new Date(2050, 11, 21), closed_at : new Date(2050, 11, 25), was_won : false };
            if (stageMilestonesByOp[opId]) {
              milestones = stageMilestonesByOp[opId];
            }
            createOpEvent_(outputCursor, opId, attendees, inviteInfo[j], false, productByOp[opId], milestones);
          }
          else {
            customer = {};
            customer[account] = 1;
            createAccountEvents_(outputCursor, attendees, customer, inviteInfo[j], productInventory);        
          }
          break;
        }
      }
    }
  
    
    if (isOverrideActive) {
      continue;
    }
      
      
    //
    // Create the event
    //
    
    if (attendeeInfo.stats.customers == 0 && attendeeInfo.stats.partners > 0) {
      createAccountEvents_(outputCursor, attendees, attendeeInfo.partners, inviteInfo[j], productInventory);
    }
    else if (attendeeInfo.stats.customer_accounts == 1) {
      
      if (IS_TRACE_ACCOUNT_ON && attendeeInfo.customers[TRACE_ACCOUNT_ID_LONG]) {
        Logger.log("TRACE Account: Found an attendee from this account in " + inviteInfo[j][SUBJECT]);
      }
    
      //Logger.log("DEBUG: found one customer account in invite: " + opInfo[j][OP_NAME]);
      
      var customerId = 0;
      let longCustomerId = 0;
      for (account in attendeeInfo.customers) {
        // The account keys for locating opportunities in the code are NOT built from 18-Digit Account IDs!
        // The account here is 18-Digit, so strip off the 3 character postfix.
        customerId = account.substring(0, account.length - 3);
        longCustomerId = account;
      }
      
      if (!customerId) {
        Logger.log("ERROR: Lost customer ID in invite: " + inviteInfo[j][SUBJECT]);
        continue;
      }
      
      // Find an opportunity
      var key = customerId + makeProductKey_(productInventory, 0);
      
      let isDefaultOp = false;
      var opId = 0;
      if (opByCustomerAndProduct[key]) {
        opId = opByCustomerAndProduct[key];
      }
      if (!opId) {
        if (numberOfOpsByCustomer[customerId] > 1) {
          isDefaultOp = true; // We were unable to deteremine the product in the invite, so selecting the default (primary) op
        }
        opId = primaryOpByCustomer[customerId]; // No product info found in invite, pick "primary" op for this period
      }
      
      if (opId) {
        let milestones = { discovery_ended : new Date(2050, 11, 21), closed_at : new Date(2050, 11, 25), was_won : false };
        if (stageMilestonesByOp[opId]) {
          milestones = stageMilestonesByOp[opId];
        }
        createOpEvent_(outputCursor, opId, attendees, inviteInfo[j], isDefaultOp, productByOp[opId], milestones);
      }
      else {
        createAccountEvents_(outputCursor, attendees, attendeeInfo.customers, inviteInfo[j], productInventory);        
      }
    }
    else if (attendeeInfo.stats.customer_accounts > 1) {
    
      if (IS_TRACE_ACCOUNT_ON && attendeeInfo.customers[TRACE_ACCOUNT_ID_LONG]) {
        Logger.log("TRACE Account: Found an attendee from this account in " + inviteInfo[j][SUBJECT]);
      }
      
      // More than one customer at this meeting. Find the one with the most representation,
      // and assume the other is there as a reference. If you can't find an op for the primary,
      // see if there is an op for the secondary. If no op, just create an event for all the
      // associated accounts.
      
      //Logger.log("DEBUG: found multiple customer accounts in invite: " + opInfo[j][OP_NAME]);
      
      // Look for at most two Customers 
      var primaryId = 0;
      var primaryCnt = 0;
      var secondaryId = 0;
      var secondaryCnt = 0;
      for (account in attendeeInfo.customers) {
        if (!primaryId || attendeeInfo.customers[account] > primaryCnt) {
          secondaryId = primaryId;
          secondaryCnt = primaryCnt;
          primaryId = account;
          primaryCnt = attendeeInfo.customers[account]; 
        }
        else if (!secondaryId || attendeeInfo.customers[account] > secondaryCnt) {
          secondaryId = account;
          secondaryCnt = attendeeInfo.customers[account];
        }
      }
      // Find an opportunity for the primary. If non, secondary.
      var productInventory = lookForProducts_(inviteInfo[j][SUBJECT] + " " + inviteInfo[j][DESCRIPTION]);
      var productKey = makeProductKey_(productInventory, 0);
   
      var opId = 0;
      let isDefaultOp = false;
      primaryId = primaryId.substring(0, primaryId.length - 3); // Strip off the extra 3 characters that make it an 18-Digit id
      if (secondaryId) {
        secondaryId = secondaryId.substring(0, secondaryId.length - 3);
      }
      if (opByCustomerAndProduct[primaryId + productKey]) {
        opId = opByCustomerAndProduct[primaryId + productKey];
      }
      else if (secondaryId && opByCustomerAndProduct[secondaryId + productKey]) {
        opId = opByCustomerAndProduct[secondaryId + productKey];
      }
      
      if (!opId) {
        if (numberOfOpsByCustomer[primaryId] > 1) {
          isDefaultOp = true; // We were unable to deteremine the product in the invite, so selecting the default (primary) op
        }
        opId = primaryOpByCustomer[primaryId]; // No product info found in invite, pick "primary" op for primary account
      }
      if (!opId) {
        if (numberOfOpsByCustomer[secondaryId] > 1) {
          isDefaultOp = true; // We were unable to deteremine the product in the invite, so selecting the default (primary) op
        }
        opId = primaryOpByCustomer[secondaryId]; // No product info found in invite, pick "primary" op for secondary account
      }
      
      if (opId) {
        let milestones = { discovery_ended : new Date(2050, 11, 21), closed_at : new Date(2050, 11, 25), was_won : false };
        if (stageMilestonesByOp[opId]) {
          milestones = stageMilestonesByOp[opId];
        }
        createOpEvent_(outputCursor, opId, attendees, inviteInfo[j], isDefaultOp, productByOp[opId], milestones);
      }
      else {
        createAccountEvents_(outputCursor, attendees, attendeeInfo.customers, inviteInfo[j], productInventory);
      }
    } 
    /*
    else {
      Logger.log("DEBUG: Unable to find an account for " + inviteInfo[j][SUBJECT]);
    }
    */
  }
}


function expand_se_events() {
  //Copy events over to Expanded tab replacing account or opportunity ids with names

  //
  // Load Opportunity Info
  //
  
  let opNameById = {}; 
  
  let sheet = SpreadsheetApp.getActive().getSheetByName(OPPORTUNITIES);
  let rangeData = sheet.getDataRange();
  let olc = rangeData.getLastColumn();
  let olr = rangeData.getLastRow();
  let scanRange = sheet.getRange(2,1, olr-1, olc);
  let opInfo = scanRange.getValues();
  
  if (olc < OP_COLUMNS) {
    Logger.log("ERROR: Imported opportunity info was only " + olc + " fields wide. Not enough! Something needs to be fixed.");
    return;
  }
  
  for (var j = 0 ; j < olr - 1; j++) {    
    opNameById[opInfo[j][OP_ID]] = opInfo[j][OP_NAME];  
  }
  
  //
  // Load Staff names
  //
  
  let staffNameById = {};
  
  sheet = SpreadsheetApp.getActive().getSheetByName(STAFF);
  rangeData = sheet.getDataRange();
  let slc = rangeData.getLastColumn();
  let slr = rangeData.getLastRow();
  scanRange = sheet.getRange(2,1, slr-1, slc);
  staffInfo = scanRange.getValues();
  
  if (slc < STAFF_COLUMNS) {
    Logger.log("ERROR: Imported Staff info was only " + slc + " fields wide. Not enough! Something's not good.");
    return;
  }
  
  // Build the user id to name mapping  
  for (var j = 0 ; j < slr - 1; j++) {
    staffNameById[staffInfo[j][STAFF_ID].trim()] = staffInfo[j][STAFF_NAME].trim();
  }
  
  //
  // Load Partner Info
  //
  
  sheet = SpreadsheetApp.getActive().getSheetByName(PARTNERS);
  rangeData = sheet.getDataRange();
  var plc = rangeData.getLastColumn();
  var plr = rangeData.getLastRow();
  scanRange = sheet.getRange(2,1, plr-1, plc);
  var partnerInfo = scanRange.getValues();
  
  if (plc < PARTNER_COLUMNS) {
    Logger.log("ERROR: Imported Partners was only " + plc + " fields wide. Not enough! Something is awry.");
    return;
  }
  
  let partnerNameById = {};
  
  for (j = 0 ; j < plr - 1; j++) {
    partnerNameById[partnerInfo[j][PARTNER_ID]]  = partnerInfo[j][PARTNER_NAME];
  }
  
  //
  // Load All Customer Info - in region first, external stuff second
  //
  
  sheet = SpreadsheetApp.getActive().getSheetByName(IN_REGION_CUSTOMERS);
  rangeData = sheet.getDataRange();
  let alc = rangeData.getLastColumn();
  let alr = rangeData.getLastRow();
  scanRange = sheet.getRange(2,1, alr-1, alc);
  let customerInfo = scanRange.getValues();
  
  if (alc < CUSTOMER_COLUMNS) {
    Logger.log("ERROR: Imported Customers was only " + alc + " fields wide. Not enough! Something is wrong.");
    return;
  }
  
  let customerNameById = {};
 
  for (j = 0 ; j < alr - 1; j++) {
    customerNameById[customerInfo[j][CUSTOMER_ID]] = customerInfo[j][CUSTOMER_NAME];
  }
  
  sheet = SpreadsheetApp.getActive().getSheetByName(EXTERNAL_CUSTOMERS);
  rangeData = sheet.getDataRange();
  alc = rangeData.getLastColumn();
  alr = rangeData.getLastRow();
  if (alr > 1) {
    
    scanRange = sheet.getRange(2,1, alr-1, alc);
    customerInfo = scanRange.getValues();
    
    if (alc < CUSTOMER_COLUMNS) {
      Logger.log("ERROR: Imported Customers was only " + alc + " fields wide. Not enough! Something is wrong.");
      return; 
    }
    
    for (j = 0 ; j < alr - 1; j++) {
      customerNameById[customerInfo[j][CUSTOMER_ID]] = customerInfo[j][CUSTOMER_NAME];
    }
  }
  
  //
  // Copy events over, replacing account or opportunity ids with names
  //

  sheet = SpreadsheetApp.getActive().getSheetByName(EVENTS);
  rangeData = sheet.getDataRange();
  let elc = rangeData.getLastColumn();
  let elr = rangeData.getLastRow();
  scanRange = sheet.getRange(1,1, elr, elc);
  let eventInfo = scanRange.getValues();
  
  if (elc < EVENT_COLUMNS) {
    Logger.log("ERROR: Imported Customers was only " + elc + " fields wide. Not enough! Something is wrong.");
    return;
  }
  
  // Set event output cursor
  sheet = SpreadsheetApp.getActive().getSheetByName(EVENTS_EXPANDED);
  sheet.clearContents();  
  let outputRange = sheet.getRange(1,1);
  let rowOffset = 0;
 
  for (j = 0 ; j < elr; j++) {
    let name = "Unknown";
    let type = "Unknown";
    if (opNameById[eventInfo[j][EVENT_RELATED_TO]]) {
      name = opNameById[eventInfo[j][EVENT_RELATED_TO]];
      type = "Opportunity";
    }
    else if (partnerNameById[eventInfo[j][EVENT_RELATED_TO]]) {
      name = partnerNameById[eventInfo[j][EVENT_RELATED_TO]];
      type = "Partner";
    }
    else if (customerNameById[eventInfo[j][EVENT_RELATED_TO]]) {
      name = customerNameById[eventInfo[j][EVENT_RELATED_TO]];
      type = "Customer";
    }
    if (0 == rowOffset) {
      // We are copying over the header
      outputRange.offset(rowOffset, 0).setValue("Event Type");
      outputRange.offset(rowOffset, 1).setValue("Name");
      outputRange.offset(rowOffset, 2).setValue("Assigned To");
    }
    else {
      outputRange.offset(rowOffset, 0).setValue(type);
      outputRange.offset(rowOffset, 1).setValue(name);  
      outputRange.offset(rowOffset, 2).setValue(staffNameById[eventInfo[j][EVENT_ASSIGNED_TO]]);
    }
    outputRange.offset(rowOffset, 3).setValue(eventInfo[j][EVENT_MEETING_TYPE]);
    outputRange.offset(rowOffset, 4).setValue(eventInfo[j][EVENT_SUBJECT]);
    outputRange.offset(rowOffset, 5).setValue(eventInfo[j][EVENT_OP_STAGE]);
    outputRange.offset(rowOffset, 6).setValue(eventInfo[j][EVENT_START]);
    outputRange.offset(rowOffset, 7).setValue(eventInfo[j][EVENT_END]);
    outputRange.offset(rowOffset, 8).setValue(eventInfo[j][EVENT_PRODUCT]);
    outputRange.offset(rowOffset, 9).setValue(eventInfo[j][EVENT_DESC]);
    outputRange.offset(rowOffset, 10).setValue(eventInfo[j][EVENT_REP_ATTENDED]);
    outputRange.offset(rowOffset, 11).setValue(eventInfo[j][EVENT_LOGISTICS]);
    outputRange.offset(rowOffset, 12).setValue(eventInfo[j][EVENT_PREP_TIME]);
    
    rowOffset++;
  }
}

function createSpecialEvents_(outputTab, attendees, inviteInfo, productInventory, meetingType) {
  // No account for these special events
  
  // Logger.log("DEBUG: Entered createSpecialEvents_ for: " + inviteInfo[SUBJECT]);
  
  let assignedTo = staffEmailToIdMap[inviteInfo[ASSIGNED_TO]];
  let descriptionScan = filterAndAnalyzeDescription_(inviteInfo[DESCRIPTION]); // returns filteredText, hasTeleconference and prepTime
  
  let logistics = "Face to Face";
  if (descriptionScan.hasTeleconference) {
    logistics = "Remote";
  }
  
  let repAttended = "No";
  if (isRepPresent_(inviteInfo[CREATED_BY], attendees)) {
    repAttended = "Yes"; 
  }
  
  
  
  // Logger.log("Debug: Creating createAccountEvents_ for: " + account);
  
  outputTab.range.offset(outputTab.rowOffset, EVENT_ASSIGNED_TO).setValue(assignedTo);
  outputTab.range.offset(outputTab.rowOffset, EVENT_OP_STAGE).setValue("None"); // None accepts all meeting types
  outputTab.range.offset(outputTab.rowOffset, EVENT_MEETING_TYPE).setValue(meetingType);
  outputTab.range.offset(outputTab.rowOffset, EVENT_SUBJECT).setValue(inviteInfo[SUBJECT]);
  outputTab.range.offset(outputTab.rowOffset, EVENT_START).setValue(inviteInfo[START]);
  outputTab.range.offset(outputTab.rowOffset, EVENT_END).setValue(inviteInfo[END]);
  outputTab.range.offset(outputTab.rowOffset, EVENT_REP_ATTENDED).setValue(repAttended);
  outputTab.range.offset(outputTab.rowOffset, EVENT_PRODUCT).setValue(getOneProduct_(productInventory));
  outputTab.range.offset(outputTab.rowOffset, EVENT_DESC).setValue(descriptionScan.filteredText + "\nProducts: " + getProducts_(productInventory) + "\nAttendees: " + inviteInfo[ATTENDEE_STR]);
  outputTab.range.offset(outputTab.rowOffset, EVENT_LOGISTICS).setValue(logistics);
  outputTab.range.offset(outputTab.rowOffset, EVENT_PREP_TIME).setValue(descriptionScan.prepTime);
  
  outputTab.rowOffset++;
  
}

function createAccountEvents_(outputTab, attendees, attendeeInfo, inviteInfo, productInventory) {

  // Logger.log("DEBUG: Entered createAccountEvents_ for: " + inviteInfo[SUBJECT]);
  
  let assignedTo = staffEmailToIdMap[inviteInfo[ASSIGNED_TO]];
  let descriptionScan = filterAndAnalyzeDescription_(inviteInfo[DESCRIPTION]); // returns filteredText, hasTeleconference and prepTime
  
  let logistics = "Face to Face";
  if (descriptionScan.hasTeleconference) {
    logistics = "Remote";
  }
  
  let repAttended = "No";
  if (isRepPresent_(inviteInfo[CREATED_BY], attendees)) {
    repAttended = "Yes"; 
  }
  
  let event = lookForMeetingType_("Discovery & Qualification", inviteInfo[SUBJECT] + " " + descriptionScan.filteredText); // There is no lead gen stage
  
  for (account in attendeeInfo) {
    
    // Logger.log("Debug: Creating createAccountEvents_ for: " + account);
    
    outputTab.range.offset(outputTab.rowOffset, EVENT_ASSIGNED_TO).setValue(assignedTo);
    outputTab.range.offset(outputTab.rowOffset, EVENT_OP_STAGE).setValue("None"); // None accepts all meeting types
    outputTab.range.offset(outputTab.rowOffset, EVENT_MEETING_TYPE).setValue(event.meeting);
    outputTab.range.offset(outputTab.rowOffset, EVENT_RELATED_TO).setValue(account); // relate directly to the account
    outputTab.range.offset(outputTab.rowOffset, EVENT_SUBJECT).setValue(inviteInfo[SUBJECT]);
    outputTab.range.offset(outputTab.rowOffset, EVENT_START).setValue(inviteInfo[START]);
    outputTab.range.offset(outputTab.rowOffset, EVENT_END).setValue(inviteInfo[END]);
    outputTab.range.offset(outputTab.rowOffset, EVENT_REP_ATTENDED).setValue(repAttended);
    outputTab.range.offset(outputTab.rowOffset, EVENT_PRODUCT).setValue(getOneProduct_(productInventory));
    outputTab.range.offset(outputTab.rowOffset, EVENT_DESC).setValue(descriptionScan.filteredText + "\nProducts: " + getProducts_(productInventory) + "\nAttendees: " + inviteInfo[ATTENDEE_STR]);
    outputTab.range.offset(outputTab.rowOffset, EVENT_LOGISTICS).setValue(logistics);
    outputTab.range.offset(outputTab.rowOffset, EVENT_PREP_TIME).setValue(descriptionScan.prepTime);
    
    outputTab.rowOffset++;
    
  }
}

function createOpEvent_(outputTab, opId, attendees, inviteInfo, isDefaultOp, opProduct, opMilestones) {
  // We are only placing an event in three phases of the Op (4 stages, Closed get's two)
  // - Discovery & Qualification
  // - Technical & Business Validation
  // - Closed (/Won or /Lost)
  //
  // All the other stages aren't relevant to SE activity (and in fact block various meeting types, so we must
  // stay away from selecting them.)
  
  if (!opMilestones) {
    Logger.log("ERROR: Lost opportunity's stage milestones!");
    return;
  }
  
  //Logger.log("DEBUG: createOpEvent_: " + inviteInfo[SUBJECT]);
  let descriptionScan = filterAndAnalyzeDescription_(inviteInfo[DESCRIPTION]);
  
  let logistics = "Face to Face";
  if (descriptionScan.hasTeleconference) {
    logistics = "Remote";
  }
  
  // Determine stage of opportunity when invite occurred
  let inviteDate = Date.parse(inviteInfo[START]);
  let opStage = "Discovery & Qualification";
  
  if (inviteDate > opMilestones.closed_at) {
    if (opMilestones.was_won == true) {
      opStage = "Closed/Won";
    }
    else {
      opStage = "Closed/Lost";
    }
  }
  else if (inviteDate > opMilestones.discovery_ended) {   
    opStage = "Technical & Business Validation";   
  }
  
  let assignedTo = staffEmailToIdMap[inviteInfo[ASSIGNED_TO]];
  
  // FIXME Only knows about South Strategic EAMs
  let repAttended = "No";
  if (isRepPresent_(inviteInfo[CREATED_BY], attendees)) {
    repAttended = "Yes"; 
  }
  
  /*
  if (inviteInfo[ASSIGNED_TO] == accountTeamByOp[opId].rep ||
  inviteInfo[ASSIGNED_TO] == accountTeamByOp[opId].rep
  repAttended = "Yes";
  */
  
  
  let event = lookForMeetingType_(opStage, inviteInfo[SUBJECT] + " " + descriptionScan.filteredText);
  
  outputTab.range.offset(outputTab.rowOffset, EVENT_ASSIGNED_TO).setValue(assignedTo);
  outputTab.range.offset(outputTab.rowOffset, EVENT_OP_STAGE).setValue(event.stage);
  outputTab.range.offset(outputTab.rowOffset, EVENT_MEETING_TYPE).setValue(event.meeting); 
  outputTab.range.offset(outputTab.rowOffset, EVENT_RELATED_TO).setValue(opId); // Relate to the opportunity
  outputTab.range.offset(outputTab.rowOffset, EVENT_SUBJECT).setValue(inviteInfo[SUBJECT]);
  outputTab.range.offset(outputTab.rowOffset, EVENT_START).setValue(inviteInfo[START]);
  outputTab.range.offset(outputTab.rowOffset, EVENT_END).setValue(inviteInfo[END]);
  outputTab.range.offset(outputTab.rowOffset, EVENT_REP_ATTENDED).setValue(repAttended);
  outputTab.range.offset(outputTab.rowOffset, EVENT_PRODUCT).setValue(opProduct);
  if (isDefaultOp) {
    outputTab.range.offset(outputTab.rowOffset, EVENT_DESC).setValue(descriptionScan.filteredText + "\nDefault Op Selected.\nAttendees: " + inviteInfo[ATTENDEE_STR]);
  }
  else {
    outputTab.range.offset(outputTab.rowOffset, EVENT_DESC).setValue(descriptionScan.filteredText + "\nAttendees: " + inviteInfo[ATTENDEE_STR]);
  }
  outputTab.range.offset(outputTab.rowOffset, EVENT_LOGISTICS).setValue(logistics); 
  outputTab.range.offset(outputTab.rowOffset, EVENT_PREP_TIME).setValue(descriptionScan.prepTime);
  
  outputTab.rowOffset++;  
  
}
