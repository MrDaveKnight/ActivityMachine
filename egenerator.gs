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
  // - The products being discussed
  // - The "general" type of activity (workshop, demo, presentation, more on this later)
  //
  // Determine the CUSTOMERS associated with all of the attendees. Determine which, if any,
  // customer has the majority of attendees. Assume this is the customer the invite is "targeted" for. Others
  // are partners or contractors.
  //
  // Determine what PRODUCTS are being discussed. Look for <product> keywords in the subject and 
  // description. Reference the lookForProducts_() function. The result will be zero, one or more products. 
  //
  // For a targeted customer invite, search for an OPPORTUNITY that matches the targeted customer and products. If
  // multiple products are in the invite, and no corresponding opportunity is found for the same set of products, look for 
  // an opportunity with one of the products. If a matching opportunity is found, create an SE activity for that 
  // opportunity. If no opportunities are found, create an SE activity for the customer's default/primary opportunity 
  // If there are no opportunities for the targeted customers, create an SE activity for
  // the account directly.
  //
  // When searching for opportunities, if two or more are found that match the customer and product, 
  // select the one with close date in closest proximity to the invite date.
  //
  // When selecting the default/primary OPPORTUNITY, chose the one that is active, with the most recent updates 
  // in Salesforce
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
  // required of an SE to do it, i.e. for easy to hard ...
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
  //  in certain Opportunity Stages, so they will invoke a "forced stage SWITCH". For the up-to-date mapping, reference
  //  the lookForMeetingType_ function in the eheuristic.gs file.
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
  
  
  //
  // Initialize global log
  //
  
  logStamp("SE Event Build");
    
  clearTab_(EVENTS, EVENT_HEADER);
  clearTab_(EVENTS_UNVEILED, REVIEW_HEADER);
  
  // Information for accounts, staff, opportunities and calendar invites
  // is loaded from tabs in the spreadsheet into two-dimentional arrays with
  // this naming convention -
  //
  //      <info-type>Info 
  //
  
  //
  // Load Account Info
  //
  
  // Load "in-play" customers
  if (!load_customer_info_(false, emailToCustomerMapG, accountTypeG, 2, true)) { 
    return;
  }
  // Loads "in-play" partners
  if (!load_partner_info_(false, emailToPartnerMapG, accountTypeG, 2, true)) {
    return;
  }
  
  load_lead_info_(true); // Always filters just what's in play (there are hundreds of thousands of leads man).
  
  createChoiceLists(); // I wants the phase 2 maps
  
  //
  // Load Staff Info - SEs and Reps
  //
  
  var sheet = SpreadsheetApp.getActive().getSheetByName(STAFF);
  var rangeData = sheet.getDataRange();
  var slc = rangeData.getLastColumn();
  var slr = rangeData.getLastRow();
  var scanRange = sheet.getRange(2,1, slr-1, slc);
  var staffInfo = scanRange.getValues();
  
  if (slc < STAFF_COLUMNS - 1) {
    logOneCol("ERROR - Imported Staff info was only " + slc + " fields wide. Not enough! Something's not good.");
    return;
  }
  
  if (slr < 2) logOneCol("WARNING - No staff found! Pleaase load the Staff tab before continuing. Thank you --The Management.");
  
  // In case we need to generate long ids
  let longIdCells = sheet.getRange(2, 1); // To initialize if we have to
  
  // Build the email to user id mapping  
  for (var j = 0 ; j < slr - 1; j++) {
    if (!staffInfo[j][STAFF_EMAIL]) {
      logOneCol("WARNING - " + staffInfo[j][STAFF_NAME] + " has no email!");
      continue;
    }
    staffNameToEmailMapG[staffInfo[j][STAFF_NAME].trim()] = staffInfo[j][STAFF_EMAIL].trim();
    
    if (!staffInfo[j][STAFF_LONG_ID]) {
      let longId = generateLongId_(staffInfo[j][STAFF_ID].trim());
      longIdCells.offset(j, STAFF_LONG_ID).setValue(longId);
      staffEmailToIdMapG[staffInfo[j][STAFF_EMAIL].trim()] = longId.trim();
    }
    else {
      staffEmailToIdMapG[staffInfo[j][STAFF_EMAIL].trim()] = staffInfo[j][STAFF_LONG_ID].trim();
    }
    
    staffEmailToRoleMapG[staffInfo[j][STAFF_EMAIL].trim()] = staffInfo[j][STAFF_ROLE].trim();
    //Logger.log("DEBUG: " + staffInfo[j][STAFF_EMAIL] + " -> " + staffInfo[j][STAFF_ID]);
    
  }
  
  //
  // Load Opportunity Info
  //
  
  let targetedAccounts = loadFilter(IN_PLAY_CUSTOMERS, FILTER_ACCOUNT_ID, true); // for filtering out only the opportunities associated with accounts in the invites
  
  sheet = SpreadsheetApp.getActive().getSheetByName(OPPORTUNITIES);
  rangeData = sheet.getDataRange();
  var olc = rangeData.getLastColumn();
  var olr = rangeData.getLastRow();  
  
  if (olc < OP_COLUMNS) {
    logOneCol("ERROR: Imported opportunity info was only " + olc + " fields wide. Not enough! Something needs to be fixed.");
    return;
  }
  
  // Build the opportunity indexes  
  //
  // There may be many opportunities for a particular customer and product (renewal, new business, services, etc);
  // We want to give priority to the opportunity with the latest activity. Pick an op that brings in 
  // new business over services and renewals
  // Also track Ops that are closed/lost. We don't want those to be the default even if they were first.
  
  let opTypeIndexedByCustomerAndProduct = {}; // For op selection priority ... DEPRECATED
  let opStageIndexedByCustomerAndProduct = {}; // For op selection priority 
  let primaryOpTypeIndexedByCustomer = {}; // For selecting the primary opportunity out of a set with different products ... DEPRECATED
  let primaryOpStageIndexedByCustomer = {}; // Don't want Closed/Lost to be a default
  let primaryOpActivityDateIndexedByCustomer = {}; // What to target op with most recent activity
  
  let chunkSize = 1000;
  let chunkFirstRow = 2;
  let chunkLastRow = chunkSize < olr ? chunkSize : olr;
  let chunkLength = chunkLastRow - chunkFirstRow + 1;
  while (chunkLength > 0) {
    
    let scanRange = sheet.getRange(chunkFirstRow,1, chunkLength, olc);
    let opInfo = scanRange.getValues();
    
    for (let j = 0; j < chunkLength; j++) {
      
      if (!targetedAccounts[opInfo[j][OP_ACCOUNT_ID]]) {
        continue;
      }
      
      let scanResults = lookForProducts_(opInfo[j][OP_NAME]);
      let productKey = makeProductKey_(scanResults, opInfo[j][OP_PRIMARY_PRODUCT]);
      if (productKey == "-") {
        logOneCol("NOTICE - " + opInfo[j][OP_NAME] + " has no product!");
      }
      
      let latestActivityDate = Date.parse(opInfo[j][OP_ACTIVITY_DATE]);
      
      // Important Note:
      // The OP_ACCOUNT_ID provided in the opportunity record does NOT have the
      // extra 3 characters at the end needed to make it a so called "18 Digit Account ID".
      // However, the account records we use to find accounts from emails DOES
      // use the full "18 Digit Account ID". Since these account IDs need to sync
      // up at some point, we will be stripping off the 3 character postfix when
      // we create the keys from emails. We don't do that here mind you (they are already gone),
      // but later we will. Just a heads up on this goofy problem created by Salesforce
      let key = opInfo[j][OP_ACCOUNT_ID] + productKey;
      
      if (!numberOfOpsByCustomerG[opInfo[j][OP_ACCOUNT_ID]]) {
        numberOfOpsByCustomerG[opInfo[j][OP_ACCOUNT_ID]] = 1;
      }
      else {
        numberOfOpsByCustomerG[opInfo[j][OP_ACCOUNT_ID]]++;
      }
      
      if (opByCustomerAndProductG[key]) {
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
        
        opByCustomerAndProductG[key] = opInfo[j][OP_ID];
        opTypeIndexedByCustomerAndProduct[key] = opInfo[j][OP_TYPE];
        }
        */
        
        // Select an active op if available
        if (opInfo[j][OP_STAGE].indexOf("Closed") == -1 && opStageIndexedByCustomerAndProduct[key].indexOf("Closed") != -1) {
          opByCustomerAndProductG[key] = opInfo[j][OP_ID];
          opTypeIndexedByCustomerAndProduct[key] = opInfo[j][OP_TYPE];
          opStageIndexedByCustomerAndProduct[key] = opInfo[j][OP_STAGE];
        }
      }
      else {
        
        
        if (IS_TRACE_ACCOUNT_ON && TRACE_ACCOUNT_ID == opInfo[j][OP_ACCOUNT_ID]) {
          Logger.log("TRACE Account: Tracking op " + opInfo[j][OP_ID] + ":" + opInfo[j][OP_TYPE] + productKey);
        }
        
        opByCustomerAndProductG[key] = opInfo[j][OP_ID];
        opTypeIndexedByCustomerAndProduct[key] = opInfo[j][OP_TYPE];
        opStageIndexedByCustomerAndProduct[key] = opInfo[j][OP_STAGE];
      }
      if (!primaryOpByCustomerG[opInfo[j][OP_ACCOUNT_ID]]) {
        if (IS_TRACE_ACCOUNT_ON && TRACE_ACCOUNT_ID == opInfo[j][OP_ACCOUNT_ID]) {
          Logger.log("TRACE Account: Tracking primary op "+ opInfo[j][OP_ID] + ":" + opInfo[j][OP_TYPE] + productKey);
        }
        primaryOpByCustomerG[opInfo[j][OP_ACCOUNT_ID]] = opInfo[j][OP_ID]; // Pick an op for all invites with no product specified to go to
        primaryOpTypeIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] = opInfo[j][OP_TYPE];
        primaryOpStageIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] = opInfo[j][OP_STAGE];
        primaryOpActivityDateIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] = latestActivityDate;    
      }
      else {
        if ((opInfo[j][OP_STAGE].indexOf("Closed") == -1 &&  primaryOpStageIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]].indexOf("Closed") != -1) || 
          (primaryOpActivityDateIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] < latestActivityDate)) {
            primaryOpByCustomerG[opInfo[j][OP_ACCOUNT_ID]] = opInfo[j][OP_ID]; // Pick an op for all invites with no product specified to go to
            primaryOpTypeIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] = opInfo[j][OP_TYPE];
            primaryOpStageIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] = opInfo[j][OP_STAGE];
            primaryOpActivityDateIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] = latestActivityDate;
          }
        
        /*
        for debug logging
        if ((opInfo[j][OP_STAGE].indexOf("Closed") == -1 &&  primaryOpStageIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]].indexOf("Closed") != -1)) {
        primaryOpByCustomerG[opInfo[j][OP_ACCOUNT_ID]] = opInfo[j][OP_ID]; // Pick an op for all invites with no product specified to go to
        primaryOpTypeIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] = opInfo[j][OP_TYPE];
        primaryOpStageIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] = opInfo[j][OP_STAGE];
        primaryOpActivityDateIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] = latestActivityDate;
        }
        else if ((primaryOpActivityDateIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] < latestActivityDate)) {
        Logger.log("Reselected a primary op. Old: " + primaryOpTypeIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] + ":" + primaryOpStageIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] + ":" + primaryOpActivityDateIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] + ", New: " + 
        opInfo[j][OP_TYPE] + ":" + opInfo[j][OP_STAGE] + ":" + latestActivityDate + "(" + opInfo[j][OP_ACTIVITY_DATE] + ")");
        primaryOpByCustomerG[opInfo[j][OP_ACCOUNT_ID]] = opInfo[j][OP_ID]; // Pick an op for all invites with no product specified to go to
        primaryOpTypeIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] = opInfo[j][OP_TYPE];
        primaryOpStageIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] = opInfo[j][OP_STAGE];
        primaryOpActivityDateIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] = latestActivityDate;
        }
        */
      }
      
      
      /* deprecated
      else if ((primaryOpStageIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] == "Closed/Lost") ||
      (primaryOpTypeIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] == "Services" && opInfo[j][OP_TYPE] != "Services") ||
      (primaryOpTypeIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] == "Renewal" && (opInfo[j][OP_TYPE] != "Services" && opInfo[j][OP_TYPE] != "Services"))) {
      // Make the primary something other than services or renewals if something else exists
      if (IS_TRACE_ACCOUNT_ON && TRACE_ACCOUNT_ID == opInfo[j][OP_ACCOUNT_ID]) {
      Logger.log("TRACE Account: Resetting primary op to "+ opInfo[j][OP_ID] + ":" + opInfo[j][OP_TYPE] + productKey);
      }
      primaryOpByCustomerG[opInfo[j][OP_ACCOUNT_ID]] = opInfo[j][OP_ID]; // Pick an op for all invites with no product specified to go to   
      primaryOpTypeIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] = opInfo[j][OP_TYPE];
      primaryOpStageIndexedByCustomer[opInfo[j][OP_ACCOUNT_ID]] = opInfo[j][OP_STAGE];
      } */
      
      primaryProductByOpG[opInfo[j][OP_ID]] = opInfo[j][OP_PRIMARY_PRODUCT];
      if (opInfo[j][OP_PRIMARY_PRODUCT] != "Terraform" && 
          opInfo[j][OP_PRIMARY_PRODUCT] != "Vault" &&
          opInfo[j][OP_PRIMARY_PRODUCT] != "Consul" &&
          opInfo[j][OP_PRIMARY_PRODUCT] != "Nomad") {
        let prod = opInfo[j][OP_PRIMARY_PRODUCT].toLowerCase();
        if (prod.indexOf("terraform") != -1 || prod.indexOf("tfe") || prod.indexOf("tfc")) {
          primaryProductByOpG[opInfo[j][OP_ID]] = "Terraform";
        }
        else if (prod.indexOf("v") != -1) {
          primaryProductByOpG[opInfo[j][OP_ID]] = "Vault";
        }
        else if (prod.indexOf("c") != -1) {
          primaryProductByOpG[opInfo[j][OP_ID]] = "Consul";
        }
        else if (prod.indexOf("n") != -1) {
          primaryProductByOpG[opInfo[j][OP_ID]] = "Nomad";
        }
        else {
          primaryProductByOpG[opInfo[j][OP_ID]] = "N/A";
          Logger.log("WARNING - Primary product for Op " + opInfo[j][OP_NAME] + " wasn't valid: " + opInfo[j][OP_PRIMARY_PRODUCT]);
        }
      }
      let team = {se_primary : opInfo[j][OP_SE_PRIMARY] , se_secondary : opInfo[j][OP_SE_SECONDARY] , rep : opInfo[j][OP_OWNER] };
      accountTeamByOpG[opInfo[j][OP_ID]] = team; // Also used as an op history fitler
      
    }
    
    chunkFirstRow = chunkLastRow + 1;
    chunkLastRow += chunkSize;
    chunkLength = chunkSize;
    if (chunkLastRow > olr) {
      chunkLastRow = olr;
      chunkLength = chunkLastRow - chunkFirstRow + 1
    }
  } 
  
  //
  // Load Opportunity Stage History
  //
  
  var sheet = SpreadsheetApp.getActive().getSheetByName(HISTORY);
  rangeData = sheet.getDataRange();
  var hlc = rangeData.getLastColumn();
  var hlr = rangeData.getLastRow(); 
  
  if (hlc < HISTORY_COLUMNS) {
    logOneCol("ERROR: Imported opportunity history was only " + hlc + " fields wide. Not enough! Qu'est-ce que c'est?");
    return;
  }
  
  
  chunkSize = 1000;
  chunkFirstRow = 2;
  chunkLastRow = chunkSize < hlr ? chunkSize : hlr;
  chunkLength = chunkLastRow - chunkFirstRow + 1;
  while (chunkLength > 0) {
    let scanRange = sheet.getRange(chunkFirstRow,1, chunkLength, hlc);
    let historyInfo = scanRange.getValues();
    
    for (let j = 0; j < chunkLength; j++) {
      
      if (accountTeamByOpG[historyInfo[j][HISTORY_OP_ID]]) {
        // This op is being tracked
        
        if (!stageMilestonesByOpG[historyInfo[j][HISTORY_OP_ID]]) {
          stageMilestonesByOpG[historyInfo[j][HISTORY_OP_ID]] = { discovery_ended : new Date(2050, 11, 21), closed_at : new Date(2050, 11, 25), was_won : false };
        }
        
        var milestone = stageMilestonesByOpG[historyInfo[j][HISTORY_OP_ID]];
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
    
    chunkFirstRow = chunkLastRow + 1;
    chunkLastRow += chunkSize;
    chunkLength = chunkSize;
    if (chunkLastRow > hlr) {
      chunkLastRow = hlr;
      chunkLength = chunkLastRow - chunkFirstRow + 1
    }
  } 
  
  //
  // Load Special meetings and Account overrides
  //
  
  let parms = SpreadsheetApp.getActive().getSheetByName(RUN_PARMS); 
  let overrideRange = parms.getRange(5,7,16,5); // Hardcoded to format of cells in RUN_PARMS!
  let overrides = overrideRange.getValues();
  let specialRange = parms.getRange(31,7,16,3); // Hardcoded to format of cells in RUN_PARMS!
  let specials = specialRange.getValues();
  let bogusRange = parms.getRange(53,7,16,2); // Hardcoded to format of cells in RUN_PARMS!
  let bogusStuff = bogusRange.getValues();
  
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
    logOneCol("ERROR: Imported Calendar was only " + lastColumn + " fields wide. Not enought! Something bad happened.");
    return;
  }
  scanRange = sheet.getRange(2,1, lastRow-1, lastColumn);
  
  // Suck all the calendar data up into memory.
  let inviteInfo = scanRange.getValues();
  
  // Set event output cursor
  sheet = SpreadsheetApp.getActive().getSheetByName("EVENTS");
  var elr = sheet.getLastRow();
  var outputRange = sheet.getRange(elr+1,1);
  var outputCursor = {range : outputRange, rowOffset : 0};
  
  // 
  // Do what we came here to do!
  // Convert calendar invites (Calendar tab) to SE events (Events tab)
  //
  
  let eventCount = 0;
  
  for (j = 0 ; j < lastRow - 1; j++) {
    
    //  
    // Look for and track the "PREP" invites. Save time in minutes
    //  
    
    let subjectLine = inviteInfo[j][SUBJECT].split(":");
    let subjectTag = subjectLine[0].trim().toLowerCase();
    let subjectTarget = "";  
    if (subjectTag == PREP && subjectLine[1]) {
      subjectTarget = subjectLine[1].trim(); // Subject line must match exactly
      let s = new Date(inviteInfo[j][START]);
      let e = new Date(inviteInfo[j][END]);
      let m = (e.getTime() - s.getTime()) / 60000; // minutes
      if (prepCalendarEntries[subjectTarget]) {
        prepCalendarEntries[subjectTarget] += m;
      }   
      else {
        prepCalendarEntries[subjectTarget] = m; 
      }      
      // Logger.log("DEBUG: Prep for " + subjectLine[1] + " is " + m);
    }    
    
    if (!inviteInfo[j][ASSIGNED_TO] || !inviteInfo[j][ATTENDEE_STR] || !inviteInfo[j][START]) {
      continue;
    }
    
    var attendees = inviteInfo[j][ATTENDEE_STR].split(","); // convert comma separated list of emails converted to an array
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
    // Look for "special" known meetings we want to track
    // 
    
    let isSpecialActive = false;
    if (inviteInfo[j][ASSIGNEE_STATUS] != "NO") {
      for (let row in specials) {
        if (specials[row][0]) {      
          
          let meetingType = specials[row][0];
          let subjectTest = false;
          let emailTest = false;
          
          if (!specials[row][1] && !specials[row][2]) {
            continue;
          }
          
          if (specials[row][1]) {
            let subjectRegex = new RegExp(specials[row][1]);
            subjectTest = subjectRegex.test(inviteInfo[j][SUBJECT])
          }
          else {
            subjectTest = true; 
          }
          
          if (specials[row][2]) {
            let emailRegex = new RegExp(specials[row][2]);
            emailTest = emailRegex.test(inviteInfo[j][ATTENDEE_STR])
          }
          else {
            emailTest = true; 
          }
          
          if (subjectTest && emailTest) {
            
            // Give priority to subject
            let pi = lookForProducts_(inviteInfo[j][SUBJECT]);
            if (pi.count == 0) {
              pi = lookForProducts_(inviteInfo[j][DESCRIPTION]);
            }
            createSpecialEvents_(outputCursor, attendees, inviteInfo[j], pi, meetingType);
            eventCount++;
            isSpecialActive = true;
            logOneCol("NOTICE - " + inviteInfo[j][SUBJECT] + " is a special meeting.");
            break;
          }
        }
      }
    }
    if (isSpecialActive) {
      continue;
    }
    
    //
    // Skip bogus meetings
    //
    
    let foundBogusMeeting = false;
    for (let row in bogusStuff) {
      if (bogusStuff[row][0] || bogusStuff[row][1]) {      
        
        let subjectTest = false;
        let emailTest = false;
        
        if (bogusStuff[row][0]) {
          let subjectRegex = new RegExp(bogusStuff[row][0]);
          subjectTest = subjectRegex.test(inviteInfo[j][SUBJECT])
        }
        else {
          subjectTest = true; 
        }
        
        if (bogusStuff[row][1]) {
          let emailRegex = new RegExp(bogusStuff[row][1]);
          emailTest = emailRegex.test(inviteInfo[j][ATTENDEE_STR])
        }
        else {
          emailTest = true; 
        }
        
        if (subjectTest && emailTest) {
          foundBogusMeeting = true;
          break;
        }
      }
    }
    if (foundBogusMeeting) {
      logOneCol(inviteInfo[j][SUBJECT] + " is a bogus meeting.");
      continue;
    }
    
    //
    // Determine who was at the meeting
    //
    
    var attendeeInfo = lookForAccounts_(attendees, emailToCustomerMapG, emailToPartnerMapG);
    // attendeeInfo.customers - Map of prospect/customer Salesforce 18-Digit account id to number of attendees
    // attendeeInfo.partners - Map of partner Salesforce 18-Digit account id to number of attendees
    // attendeeInfo.others - Map of unknown email domains to number of attendees
    // attendeeInfo.stats.customers - Number of customers in attendence
    // attendeeInfo.stats.partners - Number of partners in attendence
    // attendeeInfo.stats.hashi - Number of hashicorp attendees
    // attendeeInfo.stats.others - Number of unidentified attendees
    
    var productInventory = lookForProducts_(inviteInfo[j][SUBJECT]); // Give priority to subject
    if (productInventory.count == 0) {
      productInventory = lookForProducts_(inviteInfo[j][DESCRIPTION]);
    }
    
    //
    // Manual Overrides
    //
    
    let isOverrideActive = false;
    let overrideAccountName = "None";
    let overrideAccountId = "None";
    for (let row in overrides) {
      if (overrides[row][1]) {      
        
        overrideAccountName = overrides[row][0];
        overrideAccountId = overrides[row][1];
        let subjectTest = false;
        let emailTest = false;
        
        if (!overrides[row][2] && !overrides[row][3]) {
          continue;
        }
        
        if (overrides[row][2]) {
          let subjectRegex = new RegExp(overrides[row][2]);
          subjectTest = subjectRegex.test(inviteInfo[j][SUBJECT])
        }
        else {
          subjectTest = true; 
        }
        
        if (overrides[row][3]) {
          let emailRegex = new RegExp(overrides[row][3]);
          emailTest = emailRegex.test(inviteInfo[j][ATTENDEE_STR])
        }
        else {
          emailTest = true; 
        }
        
        if (subjectTest && emailTest) {
          
          isOverrideActive = true;
          
          if (overrides[row][4] == "Yes") {
            customer = {};
            customer[overrideAccountId] = 1;
            eventCount += createAccountEvents_(outputCursor, attendees, customer, inviteInfo[j], productInventory);      
            break;
          }
          
          // Find an opportunity 
          let shortId = overrideAccountId.substring(0, overrideAccountId.length - 3); // Opportunities reference their accounts by the short account id, so that is how the key was made in the opBy... map
          var key = shortId + makeProductKey_(productInventory, 0);
          
          var opId = 0;
          if (opByCustomerAndProductG[key]) {
            opId = opByCustomerAndProductG[key];
          }
          
          if (!opId && productInventory.count() > 1) { 
            // Couldn't find an opportunity to match the 2 or more products in the invite, i.e "-<product-code-1><product-code-2>[<product-code-n>]"
            // So, fall back to single product opportunities. See if we can find one.
            let singleKeys = makeSingleProductKeys_(productInventory);
            for (let j = 0; j < singleKeys.length; j++) {
              opId = opByCustomerAndProductG[shortId + singleKeys[j]];
              if (opId) break;          
            }
          }
          
          if (opId) {
            
            // An opportunity can only have 1 primary product, but it may cover multiple products in it's description. If it does contain
            // multiple products, don't just default to the primary. Instead, look at the product inventory from the invite and try to
            // match one of those products in the inventory to one of the opportunity's products.
            let product = primaryProductByOpG[opId];
            if (productInventory.count() > 0 && !productInventory.has(product)) { 
              // Override the primary product of the opportunity with the main product identified in the invite. 
              // It could be that there was no opportunity for the product discussed but we had to pick one anyway ... so override.
              // Or, there were multiple products in the opportunity, but the primary didn't match the one discussed ... so again, override.
              product = productInventory.getOne();
            }
            
            let milestones = { discovery_ended : new Date(2050, 11, 21), closed_at : new Date(2050, 11, 25), was_won : false };
            if (stageMilestonesByOpG[opId]) {
              milestones = stageMilestonesByOpG[opId];
            }
            createOpEvent_(outputCursor, overrideAccountId, opId, attendees, inviteInfo[j], false, product, milestones);
            eventCount++;
          }
          else {
            customer = {};
            customer[overrideAccountId] = 1;
            eventCount += createAccountEvents_(outputCursor, attendees, customer, inviteInfo[j], productInventory);        
          }
          break;
        }
      }
    }
    
    
    if (isOverrideActive) {
      logThreeCol("NOTICE - Account Override!", "For Invite: " + inviteInfo[j][SUBJECT], "Selected Account: " + overrideAccountName + " (" + overrideAccountId + ")"); 
      continue; // DAK
    }
    
    
    //
    // Create the event
    //
    
    if (attendeeInfo.stats.customers == 0 && attendeeInfo.stats.partners > 0) {
      eventCount += createAccountEvents_(outputCursor, attendees, attendeeInfo.partners, inviteInfo[j], productInventory);
    }
    else if (attendeeInfo.stats.customer_accounts == 1) {
      
      if (IS_TRACE_ACCOUNT_ON && attendeeInfo.customers[TRACE_ACCOUNT_ID_LONG]) {
        Logger.log("TRACE Account: Found an attendee from this account in " + inviteInfo[j][SUBJECT]);
      }
      
      //Logger.log("DEBUG: found one customer account in invite: " + opInfo[j][OP_NAME]);
      
      let customerId = 0;
      let longCustomerId = 0;     
      for (account in attendeeInfo.customers) {
        // The account keys for locating opportunities in the code are NOT built from 18-Digit Account IDs!
        // The account here is 18-Digit, so strip off the 3 character postfix.
        customerId = account.substring(0, account.length - 3);
        longCustomerId = account;
      }
      
      if (!customerId) {
        logOneCol("ERROR: Lost customer ID in invite: " + inviteInfo[j][SUBJECT]);
        continue;
      }
      
      // Find an opportunity
      var key = customerId + makeProductKey_(productInventory, 0);
      
      let isDefaultOp = false;
      var opId = 0;
      if (opByCustomerAndProductG[key]) {      
        opId = opByCustomerAndProductG[key];
      } 
      
      if (!opId && productInventory.count() > 1) { 
        // Couldn't find an opportunity to match the 2 or more products in the invite, i.e "-<product-code-1><product-code-2>[<product-code-n>]"
        // So, fall back to single product opportunities. See if we can find one.
        
        let singleKeys = makeSingleProductKeys_(productInventory);
        for (let j = 0; j < singleKeys.length; j++) {
          opId = opByCustomerAndProductG[customerId + singleKeys[j]];
          if (opId) break;
        }
      }
      if (!opId) {
        if (numberOfOpsByCustomerG[customerId] > 1) {
          isDefaultOp = true; // We were unable to deteremine the product in the invite, so selecting the default (primary) op
        }
        opId = primaryOpByCustomerG[customerId]; // No product info found in invite, pick "primary" op for this period
      }
      
      if (opId) {
        let milestones = { discovery_ended : new Date(2050, 11, 21), closed_at : new Date(2050, 11, 25), was_won : false };
        if (stageMilestonesByOpG[opId]) {
          milestones = stageMilestonesByOpG[opId];
        }
        
        // An opportunity can only have 1 primary product, but it may cover multiple products in it's description. If it does contain
        // multiple products, don't just default to the primary. Instead, look at the product inventory from the invite and try to
        // match one of those products in the inventory to one of the opportunity's products.
        let product = primaryProductByOpG[opId];
        if (productInventory.count() > 0 && !productInventory.has(product)) { 
          // Override the primary product of the opportunity with the main product identified in the invite. 
          // It could be that there was no opportunity for the product discussed but we had to pick one anyway ... so override.
          // Or, there were multiple products in the opportunity, but the primary didn't match the one discussed ... so again, override.
          product = productInventory.getOne();
        }
        createOpEvent_(outputCursor, longCustomerId, opId, attendees, inviteInfo[j], isDefaultOp, product, milestones);
        eventCount++;
      }
      else {
        eventCount += createAccountEvents_(outputCursor, attendees, attendeeInfo.customers, inviteInfo[j], productInventory);        
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
      let primaryId = 0;
      let longPrimaryId = 0; 
      let primaryCnt = 0;
      let secondaryId = 0;
      let longSecondaryId = 0;
      let secondaryCnt = 0;
      for (account in attendeeInfo.customers) {
        if (!longPrimaryId || attendeeInfo.customers[account] > primaryCnt) {
          longSecondaryId = longPrimaryId;
          secondaryCnt = primaryCnt;
          longPrimaryId = account;
          primaryCnt = attendeeInfo.customers[account]; 
        }
        else if (!longSecondaryId || attendeeInfo.customers[account] > secondaryCnt) {
          longSecondaryId = account;
          secondaryCnt = attendeeInfo.customers[account];
        }
      }
      // Find an opportunity for the primary. If none, secondary.
      var productInventory = lookForProducts_(inviteInfo[j][SUBJECT]); // Give priority to subject
      if (productInventory.count == 0) {
        productInventory = lookForProducts_(inviteInfo[j][DESCRIPTION]);
      }
      var productKey = makeProductKey_(productInventory, 0);
      
      var opId = 0;
      let isDefaultOp = false;
      primaryId = longPrimaryId.substring(0, longPrimaryId.length - 3); // Strip off the extra 3 characters that make it an 18-Digit id
      if (longSecondaryId) {
        secondaryId = longSecondaryId.substring(0, longSecondaryId.length - 3);
      }
      let customerId = 0;
      if (opByCustomerAndProductG[primaryId + productKey]) {
        opId = opByCustomerAndProductG[primaryId + productKey];
        customerId = longPrimaryId;
      }
      else if (secondaryId && opByCustomerAndProductG[secondaryId + productKey]) {
        opId = opByCustomerAndProductG[secondaryId + productKey];
        customerId = longSecondaryId;
      }
      
      if (!opId && productInventory.count() > 1) { 
        // Couldn't find an opportunity to match the 2 or more products in the invite, i.e "-<product-code-1><product-code-2>[<product-code-n>]"
        // So, fall back to single product opportunities. See if we can find one.
        let singleKeys = makeSingleProductKeys_(productInventory);
        for (let j = 0; j < singleKeys.length; j++) {
          opId = opByCustomerAndProductG[primaryId + singleKeys[j]];
          if (opId) {
            customerId = longPrimaryId;
            break;
          }
        }
      }
      if (!opId && productInventory.count() > 1) { 
        // Couldn't find an opportunity to match the 2 or more products in the invite, i.e "-<product-code-1><product-code-2>[<product-code-n>]"
        // So, fall back to single product opportunities. See if we can find one.
        let singleKeys = makeSingleProductKeys_(productInventory);
        for (let j = 0; j < singleKeys.length; j++) {
          opId = opByCustomerAndProductG[secondaryId + singleKeys[j]];
          if (opId) {
            customerId = longSecondaryId;
            break;
          }
        }
      }
      
      if (!opId) {
        if (numberOfOpsByCustomerG[primaryId] > 1) {
          isDefaultOp = true; // We were unable to deteremine the product in the invite, so selecting the default (primary) op
        }
        opId = primaryOpByCustomerG[primaryId]; // No product info found in invite, pick "primary" op for primary account
        customerId = longPrimaryId;
      }
      if (!opId) {
        if (numberOfOpsByCustomerG[secondaryId] > 1) {
          isDefaultOp = true; // We were unable to deteremine the product in the invite, so selecting the default (primary) op
        }
        opId = primaryOpByCustomerG[secondaryId]; // No product info found in invite, pick "primary" op for secondary account
        customerId = longSecondaryId;
      }
      
      if (opId) {
        let milestones = { discovery_ended : new Date(2050, 11, 21), closed_at : new Date(2050, 11, 25), was_won : false };
        if (stageMilestonesByOpG[opId]) {
          milestones = stageMilestonesByOpG[opId];
        }
        
        // An opportunity can only have 1 primary product, but it may cover multiple products in it's description. If it does contain
        // multiple products, don't just default to the primary. Instead, look at the product inventory from the invite and try to
        // match one of those products in the inventory to one of the opportunity's products.
        let product = primaryProductByOpG[opId];
        if (productInventory.count() > 0 && !productInventory.has(product)) { 
          // Override the primary product of the opportunity with the main product identified in the invite. 
          // It could be that there was no opportunity for the product discussed but we had to pick an op anyway ... so override.
          // Or, there were multiple products in the opportunity, but the primary didn't match the one discussed ... so again, override.
          product = productInventory.getOne();
        }
        
        createOpEvent_(outputCursor, customerId, opId, attendees, inviteInfo[j], isDefaultOp, product, milestones);
        eventCount++;
      }
      else {
        eventCount += createAccountEvents_(outputCursor, attendees, attendeeInfo.customers, inviteInfo[j], productInventory);
      }
    } 
    else if (attendeeInfo.stats.others > 0) {
      // Could not find an account or partner. Look for a lead...
      // Logger.log(inviteInfo[j][SUBJECT] + " fell through.");
      let lead = findLead_(attendees);
      if (lead) {
        createLeadEvent_(outputCursor, lead, attendees, inviteInfo[j], productInventory);
        eventCount++;
      }
      else {       
        logThreeCol("WARNING - Unable to find a customer, partner or lead for:", inviteInfo[j][SUBJECT], inviteInfo[j][ATTENDEE_STR]);
      }     
    }
  }
  
  logOneCol("Generated a total of " + eventCount + " events.");
  logOneCol("End time: " + new Date().toLocaleTimeString());
  
}


function unveil_se_events() {
  //Copy events over to Review tab replacing account or opportunity ids with names 
  
  logStamp("Unveiling Events");
    
  //
  // Load Opportunity Info
  //
  
  // OP_ACCOUT_ID field in OPPORTUNITIES is short (15 digit), so TRUNCATE the long (18 digit) account ids! The true bool is to enable truncation.
  let targetedAccounts = loadFilter(IN_PLAY_CUSTOMERS, FILTER_ACCOUNT_ID, true); 
  let opNameById = load_map_(OPPORTUNITIES, 2, OP_COLUMNS, OP_ID, OP_NAME, targetedAccounts, OP_ACCOUNT_ID, null);
  
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
  
  if (slc < STAFF_COLUMNS - 1) {
    logOneCol("ERROR: Imported Staff info was only " + slc + " fields wide. Not enough! Something's not good.");
    return;
  }
  
  let longIdCells = sheet.getRange(2,1); // To initialize if we have to
  
  // Build the user id to name mapping  
  for (var j = 0 ; j < slr - 1; j++) {
    
    if (!staffInfo[j][STAFF_LONG_ID]) {
      let longId = generateLongId_(staffInfo[j][STAFF_ID].trim());
      longIdCells.offset(j,STAFF_LONG_ID).setValue(longId);
      staffNameById[longId.trim()] = staffInfo[j][STAFF_NAME].trim();
    }
    else {
      staffNameById[staffInfo[j][STAFF_LONG_ID].trim()] = staffInfo[j][STAFF_NAME].trim();
    }
    
  }
  
  //
  // Load Partner Info
  //
  
  let targetedPartners = loadFilter(IN_PLAY_PARTNERS, FILTER_ACCOUNT_ID, false);
  let partnerNameById = load_map_(PARTNERS, 2, PARTNER_COLUMNS, PARTNER_ID, PARTNER_NAME, targetedPartners, PARTNER_ID, null);
  
  
  
  //
  // Load All Customer Info - in region first, external stuff second
  //
  
  let inPlayCustomers = loadFilter(IN_PLAY_CUSTOMERS, FILTER_ACCOUNT_ID, false);
  let customerNameById = load_map_(IN_REGION_CUSTOMERS, 2, CUSTOMER_COLUMNS, CUSTOMER_ID, CUSTOMER_NAME, inPlayCustomers, CUSTOMER_ID, null);
  
  let externalCustomers = loadFilter(MISSING_CUSTOMERS, FILTER_ACCOUNT_ID, false);
  load_map_(ALL_CUSTOMERS, 2, CUSTOMER_COLUMNS, CUSTOMER_ID, CUSTOMER_NAME, externalCustomers, CUSTOMER_ID, customerNameById);
  
  //
  // Load Leads
  //
  
  let targetedLeads = loadFilter(POTENTIAL_LEADS, FILTER_EMAIL_DOMAIN, false);
  let leadNameById = load_map_(LEADS, 2, LEAD_COLUMNS, LEAD_ID, LEAD_NAME, targetedLeads, LEAD_EMAIL, null);
  
  //
  // Copy events over, replacing account or opportunity ids with names, or lead ids with names
  //
  
  sheet = SpreadsheetApp.getActive().getSheetByName(EVENTS);
  rangeData = sheet.getDataRange();
  let elc = rangeData.getLastColumn();
  let elr = rangeData.getLastRow();
  scanRange = sheet.getRange(2,1, elr-1, elc);
  let eventInfo = scanRange.getValues();
  
  if (elc < EVENT_COLUMNS) {
    logOneCol("ERROR: Imported Events was only " + elc + " fields wide. Not enough! Something is wrong.");
    return;
  }
  
  clearTab_(EVENTS_UNVEILED, REVIEW_HEADER);
  
  // Clear any left over update color and set event output cursor
  sheet = SpreadsheetApp.getActive().getSheetByName(EVENTS_UNVEILED);
  sheet.getRange('2:1000').setBackground('#ffffff');
  
  let outputRange = sheet.getRange(2,1); // skip over the header
  let rowOffset = 0; 
  
  let cO = 0; // Count Opportunities
  let cP = 0;
  let cC = 0;
  let cL = 0;
  let cG = 0;
  
  let reviewRowWasTouchedArray = new Array(elr).fill(false);
  PropertiesService.getScriptProperties().setProperty("reviewTouches", JSON.stringify(reviewRowWasTouchedArray));
  
  for (j = 0 ; j < elr-1; j++) {
    // Note that these types are in the Data Validation for the associated field on the Review tab
    let name = "General";
    let type = "General";
    if (opNameById[eventInfo[j][EVENT_RELATED_TO]]) {
      name = opNameById[eventInfo[j][EVENT_RELATED_TO]];
      type = "Opportunity";
      cO++;
    }
    else if (partnerNameById[eventInfo[j][EVENT_RELATED_TO]]) {
      name = partnerNameById[eventInfo[j][EVENT_RELATED_TO]];
      type = "Partner";
      cP++;
    }
    else if (customerNameById[eventInfo[j][EVENT_RELATED_TO]]) {
      name = customerNameById[eventInfo[j][EVENT_RELATED_TO]];
      type = "Customer";
      cC++;
    } 
    else if (leadNameById[eventInfo[j][EVENT_LEAD]]) {
      name = leadNameById[eventInfo[j][EVENT_LEAD]];
      type = "Lead";
      cL++;
    }
    else {
      cG++;
    }
    outputRange.offset(rowOffset, REVIEW_ASSIGNED_TO).setValue(staffNameById[eventInfo[j][EVENT_ASSIGNED_TO]]);
    outputRange.offset(rowOffset, REVIEW_EVENT_TYPE).setValue(type);
    outputRange.offset(rowOffset, REVIEW_ORIG_EVENT_TYPE).setValue(type);
    if (type == "Opportunity" || type == "Customer" || type == "Partner") {
      outputRange.offset(rowOffset, REVIEW_RELATED_TO).setValue(name); 
      outputRange.offset(rowOffset, REVIEW_ORIG_RELATED_TO).setValue(name); 
      initValidation_(outputRange.offset(rowOffset, REVIEW_RELATED_TO), type);
    }
    else if (type == "Lead") {
      outputRange.offset(rowOffset, REVIEW_LEAD).setValue(name);
      outputRange.offset(rowOffset, REVIEW_ORIG_LEAD).setValue(name);
      initValidation_(outputRange.offset(rowOffset, REVIEW_LEAD), type);
    }     
    outputRange.offset(rowOffset, REVIEW_OP_STAGE).setValue(eventInfo[j][EVENT_OP_STAGE]);
    outputRange.offset(rowOffset, REVIEW_ORIG_OP_STAGE).setValue(eventInfo[j][EVENT_OP_STAGE]);
    outputRange.offset(rowOffset, REVIEW_START).setValue(eventInfo[j][EVENT_START]);
    outputRange.offset(rowOffset, REVIEW_END).setValue(eventInfo[j][EVENT_END]);
    outputRange.offset(rowOffset, REVIEW_SUBJECT).setValue(eventInfo[j][EVENT_SUBJECT]);
    outputRange.offset(rowOffset, REVIEW_PRODUCT).setValue(eventInfo[j][EVENT_PRODUCT]);
    outputRange.offset(rowOffset, REVIEW_ORIG_PRODUCT).setValue(eventInfo[j][EVENT_PRODUCT]);
    outputRange.offset(rowOffset, REVIEW_DESC).setValue(eventInfo[j][EVENT_DESC]);
    outputRange.offset(rowOffset, REVIEW_ORIG_DESC).setValue(eventInfo[j][EVENT_DESC]);
    outputRange.offset(rowOffset, REVIEW_MEETING_TYPE).setValue(eventInfo[j][EVENT_MEETING_TYPE]);
    outputRange.offset(rowOffset, REVIEW_ORIG_MEETING_TYPE).setValue(eventInfo[j][EVENT_MEETING_TYPE]);
    outputRange.offset(rowOffset, REVIEW_REP_ATTENDED).setValue(eventInfo[j][EVENT_REP_ATTENDED]);
    outputRange.offset(rowOffset, REVIEW_ORIG_REP_ATTENDED).setValue(eventInfo[j][EVENT_REP_ATTENDED]);
    outputRange.offset(rowOffset, REVIEW_LOGISTICS).setValue(eventInfo[j][EVENT_LOGISTICS]);
    outputRange.offset(rowOffset, REVIEW_ORIG_LOGISTICS).setValue(eventInfo[j][EVENT_LOGISTICS]);
    outputRange.offset(rowOffset, REVIEW_PREP_TIME).setValue(eventInfo[j][EVENT_PREP_TIME]);
    outputRange.offset(rowOffset, REVIEW_ORIG_PREP_TIME).setValue(eventInfo[j][EVENT_PREP_TIME]);
    if (eventInfo[j][EVENT_QUALITY] == "") {
      outputRange.offset(rowOffset, REVIEW_QUALITY).setValue("Undefined");
      outputRange.offset(rowOffset, REVIEW_ORIG_QUALITY).setValue("Undefined");
    }
    else {
      outputRange.offset(rowOffset, REVIEW_QUALITY).setValue(eventInfo[j][EVENT_QUALITY]);
      outputRange.offset(rowOffset, REVIEW_ORIG_QUALITY).setValue(eventInfo[j][EVENT_QUALITY]);
    }
    outputRange.offset(rowOffset, REVIEW_NOTES).setValue(eventInfo[j][EVENT_NOTES]);
    outputRange.offset(rowOffset, REVIEW_ORIG_NOTES).setValue(eventInfo[j][EVENT_NOTES]);
    // If we ever want to upload in-region vs out-of-region meetings, we can use account type
    //outputRange.offset(rowOffset, REVIEW_ACCOUNT_TYPE).setValue(eventInfo[j][EVENT_ACCOUNT_TYPE]);
    //outputRange.offset(rowOffset, REVIEW_ORIG_ACCOUNT_TYPE).setValue(eventInfo[j][EVENT_ACCOUNT_TYPE]);   
    outputRange.offset(rowOffset, REVIEW_PROCESS).setValue(eventInfo[j][EVENT_PROCESS]);
    outputRange.offset(rowOffset, REVIEW_ORIG_PROCESS).setValue(eventInfo[j][EVENT_PROCESS]);
    
    
    rowOffset++;
  }
  
  // 
  // WARNING - This protection is hardcoded to the column format of unveiled events tab!
  // If the header changes, you may need to update this
  //
  
  var protection = sheet.protect().setDescription('Review data protection');
  var unprotected1 = sheet.getRange(2, 2, rowOffset, 3); 
  var unprotected2 = sheet.getRange(2, 8, rowOffset, 11); 
  protection.setUnprotectedRanges([unprotected1, unprotected2]);
  // Ensure the current user is an editor before removing others. Otherwise, if the user's edit
  // permission comes from a group, the script will throw an exception upon removing the group.
  var me = Session.getEffectiveUser();
  protection.addEditor(me);
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
  
  logOneCol("Unveiled " + cO + " Opportunity events.");
  logOneCol("Unveiled " + cP + " Partner events.");
  logOneCol("Unveiled " + cC + " Customer events.");
  logOneCol("Unveiled " + cL + " Lead events.");
  logOneCol("Unveiled " + cG + " General/Multi-customer events.");
  logOneCol("Unveiled a total of " + rowOffset + " events.");
  logOneCol("End time: " + new Date().toLocaleTimeString());
  
  sheet.activate();
  
}

function reconcile_se_events() {
  
  logStamp("Event Reconciliation");
  
  //
  // Load Opportunity Info
  //
  
  // OP_ACCOUT_ID field in OPPORTUNITIES is short (15 digit), so TRUNCATE the long (18 digit) account ids! The true bool is to enable truncation.
  let targetedAccounts = loadFilter(CHOICE_ACCOUNT, CHOICE_ACCOUNT_ID, true); 
  let opIdByName = load_map_(OPPORTUNITIES, 2, OP_COLUMNS, OP_NAME, OP_ID, targetedAccounts, OP_ACCOUNT_ID, null);
  
  //
  // Load Partner Info
  //
  
  let targetedPartners = loadFilter(CHOICE_PARTNER, CHOICE_PARTNER_ID, false);
  let partnerIdByName = load_map_(PARTNERS, 2, PARTNER_COLUMNS, PARTNER_NAME, PARTNER_ID, targetedPartners, PARTNER_ID, null);
  
  //
  // Load All Customer Info - in region first, external stuff second
  //
  
  let inPlayCustomers = loadFilter(CHOICE_ACCOUNT, CHOICE_ACCOUNT_ID, false);  
  let customerIdByName = load_map_(ALL_CUSTOMERS, 2, CUSTOMER_COLUMNS, CUSTOMER_NAME, CUSTOMER_ID, inPlayCustomers, CUSTOMER_ID, null);
  
  //
  // Load Leads
  //
   
  let targetedLeads = loadFilter(POTENTIAL_LEADS, FILTER_EMAIL_DOMAIN, false);
  let leadIdByName = load_map_(LEADS, 2, LEAD_COLUMNS, LEAD_NAME, LEAD_ID, targetedLeads, LEAD_EMAIL, null); 
  
  //
  // Scan unveiled events in the Review tab, updating fields in the corresponding Event tab record when necessary
  //
  
  let reviewInfo = load_tab_(EVENTS_UNVEILED, 2, REVIEW_COLUMNS);
  let eventInfo = load_tab_(EVENTS, 2, EVENT_COLUMNS);
  let reviewRowWasTouchedArray = JSON.parse(PropertiesService.getScriptProperties().getProperty("reviewTouches"));
  
  // Clear any left over update color and set event output cursor
  let sheet = SpreadsheetApp.getActive().getSheetByName(EVENTS);
  let outputRange = sheet.getRange(2,1); // skip over the header
  
  for (j = 0 ; j < reviewInfo.length; j++) {
    
    if (!reviewRowWasTouchedArray[j+2]) continue; // Array is indexed on table row number
    
    
    let relatedTo = null;
    let lead = null;
    let wasDeleted = false;
    let updatedFields = null;
    switch (reviewInfo[j][REVIEW_EVENT_TYPE]) {
      case "Opportunity":
        relatedTo = opIdByName[reviewInfo[j][REVIEW_RELATED_TO]];
        break;
      case "Partner":
        relatedTo = partnerIdByName[reviewInfo[j][REVIEW_RELATED_TO]];
        break;
      case "Customer":
        relatedTo = customerIdByName[reviewInfo[j][REVIEW_RELATED_TO]];
        break;
      case "Lead":
        lead = leadIdByName[reviewInfo[j][REVIEW_LEAD]];
        break;
      case "General":
        break;
      default:
        wasDeleted = true; // Empty cell means they deleted this row (or the field is jacked), so skip
    }
    
    if (relatedTo) {
      if (eventInfo[j][EVENT_RELATED_TO] != relatedTo) {
        if (updatedFields) {
          updatedFields += ", Related To";
        }
        else {
          updatedFields = "Related To";
        }
        outputRange.offset(j, EVENT_RELATED_TO).setValue(relatedTo);
        
        if (eventInfo[j][EVENT_LEAD] != "") {        
          updatedFields += ", Lead";
          outputRange.offset(j, EVENT_LEAD).setValue("");
        }
      }
    }
    else if (lead) {
      if (eventInfo[j][EVENT_LEAD] != lead) {
        if (updatedFields) {
          updatedFields += ", Lead";
        }
        else {
          updatedFields = "Lead";
        }
        outputRange.offset(j, EVENT_LEAD).setValue(lead);
        if (eventInfo[j][EVENT_RELATED_TO] != "") {   
          updatedFields += ", Related To";       
          outputRange.offset(j, EVENT_RELATED_TO).setValue("");
        }
      }
    } 
    if (eventInfo[j][EVENT_OP_STAGE] != reviewInfo[j][REVIEW_OP_STAGE]) {
      if (updatedFields) {
        updatedFields += ", Opportunity Stage";
      }
      else {
        updatedFields = "Opportunity Stage";
      }
      // Given the selected meeting type, we need to make sure the stage selected is allowed 
      // by Salesforce's field validation logic. We are just going to override it (no need 
      // to flag it on the UI, at least not yet.)
      let mt = reviewInfo[j][REVIEW_MEETING_TYPE];
      let s = reviewInfo[j][REVIEW_OP_STAGE];
      if (!validStagesByMeeting[mt + s]) {
        s = defaultStageForMeeting[mt]; 
        logFourCol("Overriding stage selection for API validation", "On: " +  eventInfo[j][EVENT_SUBJECT], "From: " +  reviewInfo[j][REVIEW_OP_STAGE] + ", To: " + s, "Meeting Type: " + mt);
      }
      outputRange.offset(j, EVENT_OP_STAGE).setValue(s);
    }
    if (eventInfo[j][EVENT_PRODUCT] != reviewInfo[j][REVIEW_PRODUCT]) {
      if (updatedFields) {
        updatedFields += ", Primary Product";
      }
      else {
        updatedFields = "Primary Product";
      }
      outputRange.offset(j, EVENT_PRODUCT).setValue(reviewInfo[j][REVIEW_PRODUCT]);
    }
    if (eventInfo[j][EVENT_DESC] != reviewInfo[j][REVIEW_DESC]) {
      if (updatedFields) {
        updatedFields += ", Description";
      }
      else {
        updatedFields = "Description";
      }
      outputRange.offset(j, EVENT_DESC).setValue(reviewInfo[j][REVIEW_DESC]);
    }
    if (eventInfo[j][EVENT_MEETING_TYPE] != reviewInfo[j][REVIEW_MEETING_TYPE]) {
      if (updatedFields) {
        updatedFields += ", Meeting Type";
      }
      else {
        updatedFields = "Meeting Type";
      }
      outputRange.offset(j, EVENT_MEETING_TYPE).setValue(reviewInfo[j][REVIEW_MEETING_TYPE]);
    }
    if (eventInfo[j][EVENT_REP_ATTENDED] != reviewInfo[j][REVIEW_REP_ATTENDED]) {
      if (updatedFields) {
        updatedFields += ", Rep Attended";
      }
      else {
        updatedFields = "Rep Attended";
      }
      outputRange.offset(j, EVENT_REP_ATTENDED).setValue(reviewInfo[j][REVIEW_REP_ATTENDED]);
    }    
    if (eventInfo[j][EVENT_LOGISTICS] != reviewInfo[j][REVIEW_LOGISTICS]) {
      if (updatedFields) {
        updatedFields += ", Logistics";
      }
      else {
        updatedFields = "Logistics";
      }
      outputRange.offset(j, EVENT_LOGISTICS).setValue(reviewInfo[j][REVIEW_LOGISTICS]);
    }
    if (eventInfo[j][EVENT_PREP_TIME] != reviewInfo[j][REVIEW_PREP_TIME]) {
      if (updatedFields) {
        updatedFields += ", Prep";
      }
      else {
        updatedFields = "Prep";
      }
      outputRange.offset(j, EVENT_PREP_TIME).setValue(reviewInfo[j][REVIEW_PREP_TIME]);
    }    
    let quality = reviewInfo[j][REVIEW_QUALITY];
    if (quality == "Undefined") {
      quality = "";
    }
    if (eventInfo[j][EVENT_QUALITY] != quality) {
      if (updatedFields) {
        updatedFields += ", Quality";
      }
      else {
        updatedFields = "Quality";
      }
      outputRange.offset(j, EVENT_QUALITY).setValue(quality);
    }
    if (eventInfo[j][EVENT_NOTES] != reviewInfo[j][REVIEW_NOTES]) {
      if (updatedFields) {
        updatedFields += ", Notes";
      }
      else {
        updatedFields = "Notes";
      }
      outputRange.offset(j, EVENT_NOTES).setValue(reviewInfo[j][REVIEW_NOTES]);
    }
    if (eventInfo[j][EVENT_PROCESS] != reviewInfo[j][REVIEW_PROCESS]) {
      if (updatedFields) {
        updatedFields += ", Process";
      }
      else {
        updatedFields = "Process";
      }
      outputRange.offset(j, EVENT_PROCESS).setValue(reviewInfo[j][REVIEW_PROCESS]);
    }    
    
    if (updatedFields) {
      let date = Utilities.formatDate(new Date(eventInfo[j][EVENT_START]), "GMT-5", "MMM dd, yyyy");
      logThreeCol("Record Update", eventInfo[j][EVENT_SUBJECT] + " / " + date, updatedFields);
    }
  }
}


function createSpecialEvents_(outputTab, attendees, inviteInfo, productInventory, meetingType) {
  // There is no account for these special events
  
  // Logger.log("DEBUG: Entered createSpecialEvents_ for: " + inviteInfo[SUBJECT]);
  try {
    
    let assignedTo = staffEmailToIdMapG[inviteInfo[ASSIGNED_TO]];
    let descriptionScan = filterAndAnalyzeDescription_(inviteInfo[DESCRIPTION]); // returns filteredText, hasTeleconference and prepTime
    let subjectScan = analyzeSubject_(inviteInfo[SUBJECT]);
    
    let prep = descriptionScan.prepTime; // OVERRIDES other prep calendar entries
    if (!prep) {
      prep = subjectScan.prepTime; 
    }
    
    let logistics = "Face to Face";
    if (descriptionScan.hasTeleconference || isLocationRemote_(inviteInfo[LOCATION])) {
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
    outputTab.range.offset(outputTab.rowOffset, EVENT_PREP_TIME).setValue(prep);
    outputTab.range.offset(outputTab.rowOffset, EVENT_QUALITY).setValue(descriptionScan.quality);
    outputTab.range.offset(outputTab.rowOffset, EVENT_NOTES).setValue(descriptionScan.notes);
    outputTab.range.offset(outputTab.rowOffset, EVENT_ACCOUNT_TYPE).setValue("None");
    outputTab.range.offset(outputTab.rowOffset, EVENT_PROCESS).setValue(PROCESS_UPLOAD);
    
    outputTab.rowOffset++;
    
  }
  catch (e) {
    logOneCol("ERROR - createSpecialEvents_ exception: " + e);
  }
  
}

function createAccountEvents_(outputTab, attendees, attendeeLog, inviteInfo, productInventory) {

  // Logger.log("DEBUG: Entered createAccountEvents_ for: " + inviteInfo[SUBJECT]);
  
  let eventCount = 0;
  try {
    
    let assignedTo = staffEmailToIdMapG[inviteInfo[ASSIGNED_TO]];
    let descriptionScan = filterAndAnalyzeDescription_(inviteInfo[DESCRIPTION]); // returns filteredText, hasTeleconference and prepTime
    let subjectScan = analyzeSubject_(inviteInfo[SUBJECT]);
    
    let prep = descriptionScan.prepTime; // OVERRIDES other prep calendar entries
    if (!prep) {
      prep = subjectScan.prepTime; 
    }
    
    let logistics = "Face to Face";
    if (descriptionScan.hasTeleconference || isLocationRemote_(inviteInfo[LOCATION])) {
      logistics = "Remote";
    }
    
    let repAttended = "No";
    if (isRepPresent_(inviteInfo[CREATED_BY], attendees)) {
      repAttended = "Yes"; 
    }
    
    let event = lookForMeetingType_("Discovery & Qualification", inviteInfo[SUBJECT] + " " + descriptionScan.filteredText); // There is no lead gen stage
    
    for (account in attendeeLog) {
      
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
      outputTab.range.offset(outputTab.rowOffset, EVENT_PREP_TIME).setValue(prep);
      outputTab.range.offset(outputTab.rowOffset, EVENT_QUALITY).setValue(descriptionScan.quality);
      outputTab.range.offset(outputTab.rowOffset, EVENT_NOTES).setValue(descriptionScan.notes);
      outputTab.range.offset(outputTab.rowOffset, EVENT_ACCOUNT_TYPE).setValue(accountTypeNamesG[accountTypeG[account]]);
      outputTab.range.offset(outputTab.rowOffset, EVENT_PROCESS).setValue(PROCESS_UPLOAD);
      
      outputTab.rowOffset++;
      eventCount++;
      
    }
  }
  catch (e) {
    logOneCol("ERROR - createAccountEvents_ exception: " + e);
  }
  return eventCount;
}

function createLeadEvent_(outputTab, lead, attendees, inviteInfo, productInventory) {

  try {
    
    let assignedTo = staffEmailToIdMapG[inviteInfo[ASSIGNED_TO]];
    let descriptionScan = filterAndAnalyzeDescription_(inviteInfo[DESCRIPTION]); // returns filteredText, hasTeleconference and prepTime
    let subjectScan = analyzeSubject_(inviteInfo[SUBJECT]);
    
    let prep = descriptionScan.prepTime; // OVERRIDES other prep calendar entries
    if (!prep) {
      prep = subjectScan.prepTime; 
    }
    
    let logistics = "Face to Face";
    if (descriptionScan.hasTeleconference || isLocationRemote_(inviteInfo[LOCATION])) {
      logistics = "Remote";
    }
    
    let repAttended = "No";
    if (isRepPresent_(inviteInfo[CREATED_BY], attendees)) {
      repAttended = "Yes"; 
    }
    
    let event = lookForMeetingType_("Discovery & Qualification", inviteInfo[SUBJECT] + " " + descriptionScan.filteredText); // There is no lead gen stage
    if (event.meeting != "Happy Hour" && event.meeting != "Demo") { // "Happy Hour", (short) "Demo" or "Discovery" are acceptable for leads. We better not be doing anything more 
      event.meeting = "Discovery";
    }
    
    
    //Logger.log("Debug: Running createLeadEvent_ for: " + lead);
    
    outputTab.range.offset(outputTab.rowOffset, EVENT_ASSIGNED_TO).setValue(assignedTo);
    outputTab.range.offset(outputTab.rowOffset, EVENT_OP_STAGE).setValue("None"); // None accepts all meeting types
    outputTab.range.offset(outputTab.rowOffset, EVENT_MEETING_TYPE).setValue(event.meeting);
    // outputTab.range.offset(outputTab.rowOffset, EVENT_RELATED_TO).setValue(); No accounts available to relate to - using a Lead
    outputTab.range.offset(outputTab.rowOffset, EVENT_SUBJECT).setValue(inviteInfo[SUBJECT]);
    outputTab.range.offset(outputTab.rowOffset, EVENT_START).setValue(inviteInfo[START]);
    outputTab.range.offset(outputTab.rowOffset, EVENT_END).setValue(inviteInfo[END]);
    outputTab.range.offset(outputTab.rowOffset, EVENT_REP_ATTENDED).setValue(repAttended);
    outputTab.range.offset(outputTab.rowOffset, EVENT_PRODUCT).setValue(getOneProduct_(productInventory));
    outputTab.range.offset(outputTab.rowOffset, EVENT_DESC).setValue(descriptionScan.filteredText + "\nProducts: " + getProducts_(productInventory) + "\nAttendees: " + inviteInfo[ATTENDEE_STR]);
    outputTab.range.offset(outputTab.rowOffset, EVENT_LOGISTICS).setValue(logistics);
    outputTab.range.offset(outputTab.rowOffset, EVENT_PREP_TIME).setValue(prep);
    outputTab.range.offset(outputTab.rowOffset, EVENT_QUALITY).setValue(descriptionScan.quality);
    outputTab.range.offset(outputTab.rowOffset, EVENT_NOTES).setValue(descriptionScan.notes);
    outputTab.range.offset(outputTab.rowOffset, EVENT_LEAD).setValue(lead);
    outputTab.range.offset(outputTab.rowOffset, EVENT_ACCOUNT_TYPE).setValue("Lead");
    outputTab.range.offset(outputTab.rowOffset, EVENT_PROCESS).setValue(PROCESS_UPLOAD);
    
    
    outputTab.rowOffset++;
  }
  catch (e) {
    logOneCol("ERROR - createLeadEvent_ exception: " + e);
  }
}

function createOpEvent_(outputTab, accountId, opId, attendees, inviteInfo, isDefaultOp, primaryProduct, opMilestones) {
  // We are only placing an event in three phases of the Op (4 stages, Closed get's two)
  // - Discovery & Qualification
  // - Technical & Business Validation
  // - Closed (/Won or /Lost)
  //
  // All the other stages aren't relevant to SE activity (and in fact block various meeting types, so we must
  // stay away from selecting them.)
  
  try {
   
    if (!opMilestones) {
      Logger.log("ERROR - Lost opportunity's stage milestones!");
      return;
    }
    
    //Logger.log("DEBUG: createOpEvent_: " + inviteInfo[SUBJECT]);
    let descriptionScan = filterAndAnalyzeDescription_(inviteInfo[DESCRIPTION]);
    let subjectScan = analyzeSubject_(inviteInfo[SUBJECT]);
    
    let prep = descriptionScan.prepTime; // OVERRIDES other prep calendar entries
    if (!prep) {
      prep = subjectScan.prepTime; 
    }
    
    let logistics = "Face to Face";
    if (descriptionScan.hasTeleconference || isLocationRemote_(inviteInfo[LOCATION])) {
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
    
    let assignedTo = staffEmailToIdMapG[inviteInfo[ASSIGNED_TO]];
    
    // FIXME Only knows about in-region EAMs
    let repAttended = "No";
    if (isRepPresent_(inviteInfo[CREATED_BY], attendees)) {
      repAttended = "Yes"; 
    }
    
    /*
    if (inviteInfo[ASSIGNED_TO] == accountTeamByOpG[opId].rep ||
    inviteInfo[ASSIGNED_TO] == accountTeamByOpG[opId].rep
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
    outputTab.range.offset(outputTab.rowOffset, EVENT_PRODUCT).setValue(primaryProduct);
    if (isDefaultOp) {
      outputTab.range.offset(outputTab.rowOffset, EVENT_DESC).setValue(descriptionScan.filteredText + "\nDefault Op Selected.\nAttendees: " + inviteInfo[ATTENDEE_STR]);
    }
    else {
      outputTab.range.offset(outputTab.rowOffset, EVENT_DESC).setValue(descriptionScan.filteredText + "\nAttendees: " + inviteInfo[ATTENDEE_STR]);
    }
    outputTab.range.offset(outputTab.rowOffset, EVENT_LOGISTICS).setValue(logistics); 
    outputTab.range.offset(outputTab.rowOffset, EVENT_PREP_TIME).setValue(prep);
    outputTab.range.offset(outputTab.rowOffset, EVENT_QUALITY).setValue(descriptionScan.quality);
    outputTab.range.offset(outputTab.rowOffset, EVENT_NOTES).setValue(descriptionScan.notes);
    outputTab.range.offset(outputTab.rowOffset, EVENT_ACCOUNT_TYPE).setValue(accountTypeNamesG[accountTypeG[accountId]]);
    outputTab.range.offset(outputTab.rowOffset, EVENT_PROCESS).setValue(PROCESS_UPLOAD);
    
    outputTab.rowOffset++;  
    
  }
  catch (e) {
    logOneCol("ERROR - createOpEvent_ exception: " + e);
  }  
}
