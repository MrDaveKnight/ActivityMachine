function markRunStart_() {
  PropertiesService.getScriptProperties().setProperty('whoIsRunning', Session.getActiveUser().getEmail());
}

function markRunEnd_() {
  PropertiesService.getScriptProperties().deleteProperty('whoIsRunning');
}

function dependenciesGood() {
  
  let sheet = SpreadsheetApp.getActive().getSheetByName(CHOICE_LEAD);
  
  if (!sheet) { 
    SpreadsheetApp.getUi().alert("ABORTING!\n\nPlease create a new tab called \"Lead Choices\", hide it and run again.\n\n" +
                                 "Processing introduced in 1.2.0 introduced that hidden tab to hold a list that limits the lead choices in the review tab to only valid " +
                                 "leads from the set of all attendees currently imported.");
    return false;
  }
  return true;
}

function needToAbortRun_(ui) {
  // Are we currently executing code? Used to disable menu items during runs
  if (PropertiesService.getScriptProperties().getProperty('whoIsRunning') === null) {
    return false;
  }
  else {
    let result = ui.alert("Warning!", "Previous command executed by " + PropertiesService.getScriptProperties().getProperty('whoIsRunning') +
      " is still running (or was cancelled?)\n\nDo you want to force this execution?", ui.ButtonSet.YES_NO);
    if (result == ui.Button.YES) {
      return false;
    }
  }
  return true;
}

function running_() {
  
  let driver = PropertiesService.getScriptProperties().getProperty('whoIsRunning');
  return (driver !== null)
}

function menuItem1_() {
  
  let ui = SpreadsheetApp.getUi();
  
  if (needToAbortRun_(ui)) {
    return;
  }
  
  /*
  let result = ui.alert('Please confirm', 
  'Are you sure you want to import calendar invites? This will clear the current set in the Calendar tab, and replace with a fresh import.\n\nContinue?',
  ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {
  */
  try {
    markRunStart_();
    import_calendars();
  }
  catch (e) {
    ui.alert("ERROR - unhandled exception: " + e);
  }
  finally {
    markRunEnd_();
  }
  // }
}

function menuItem2_() {
  
  let ui = SpreadsheetApp.getUi();
  
  if (needToAbortRun_(ui) || !dependenciesGood()) {
    return;
  }
  
  
  let result = ui.alert('Please confirm', 
                        'Are you sure you want to transform invites into events? This will clear the current set of Events and Expaned/Reviewed events.\n\nContinue?',
                        ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {  
    try {
      markRunStart_();
      createDataLoadFilters();
      build_se_events();
    }
    catch (e) {
      ui.alert("ERROR - unhandled exception: " + e);
    }
    finally {
      markRunEnd_();
    }
  }
}

function menuItem3_() {
  
  let ui = SpreadsheetApp.getUi();
  
  if (needToAbortRun_(ui) || !dependenciesGood()) {
    return;
  }
  
  let result = ui.alert('Please confirm', 
                        'You want to run all pre-upload event processing? This will execute the following stages in order...\n\n1. Import Calendars\n2. Generate Events\n3. Unveil Events.' +
                        '\n\nThis will clear out all invite and event data including the Review tab!\n\nAre you sure you want to do all that?',
                        ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {
    try {
      markRunStart_();
      import_calendars();
      createDataLoadFilters();
      build_se_events();
      unveil_se_events();
    }
    catch (e) {
      ui.alert("ERROR - unhandled exception: " + e);
    }
    finally {
      markRunEnd_();
    }
  }
}

function menuItem4_() {
  
  let ui = SpreadsheetApp.getUi();
  if (needToAbortRun_(ui) || !dependenciesGood()) {
    return;
  }
  
  let result = ui.alert('Please confirm', 
                        'Are you sure you want to stage events for Zapier upload? Zapier will be sending the events staged in the Upload tab to Salesforce. Once in Salesforce, the only way to delete them is 1 by 1.\n\nYou sure you want to ship these events off to Salesforce?',
                        ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {
    try {
      markRunStart_();
      stage_events_to_upload_tab();
    }
    catch (e) {
      ui.alert("ERROR - unhandled exception: " + e);
    }
    finally {
      markRunEnd_();
    }
  }
}

function menuItem6_() {
  
  let ui = SpreadsheetApp.getUi();
  
  if (needToAbortRun_(ui)) {
    return;
  }
  
  let result = ui.alert('Please confirm clear', 
                        'The Zapier "Upload SE Events" zap must be OFF before you continue!\n\nIs Zapier off?',
                        ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {
    try {
      markRunStart_();
      clearTab_(UPLOAD_STAGE, UPLOAD_STAGE_HEADER);
    }
    catch (e) {
      ui.alert("ERROR - unhandled exception: " + e);
    }
    finally {
      markRunEnd_();
    }
  }
}

function menuItem7_() {
  
  let ui = SpreadsheetApp.getUi();
  
  if (needToAbortRun_(ui) || !dependenciesGood()) {
    return;
  }
  
  let result = ui.alert('Please confirm', 
                        'You want to run Generate Events and Unveil Events?' +
                        '\n\nThis will clear out all event data including the Review tab!\n\nAre you sure you want to do all that?',
                        ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {
    try {
      markRunStart_();
      createDataLoadFilters();
      build_se_events();
      unveil_se_events();
    }
    catch (e) {
      ui.alert("ERROR - unhandled exception: " + e);
    }
    finally {
      markRunEnd_();
    }
  }
}

function menuItem8_() {
  
  let ui = SpreadsheetApp.getUi();
  
  if (needToAbortRun_(ui)) {
    return;
  }
  
  /*
  let result = ui.alert('Please confirm', 
  'You want to clear the Calendar import tab?',
  ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {
  */
  try {
    markRunStart_();
    clearTab_(CALENDAR, CALENDAR_HEADER);
  }
  catch (e) {
    ui.alert("ERROR - unhandled exception: " + e);
  }
  finally {
    markRunEnd_();
  }
  //}
}

function menuItem9_() {
  
  let ui = SpreadsheetApp.getUi();
  
  if (needToAbortRun_(ui)) {
    return;
  }
  
  /*
  let result = ui.alert('Please confirm', 
  'You want to clear the Event staging tab?',
  ui.ButtonSet.YES_NO);
  
  
  if (result == ui.Button.YES) {
  */
  try {
    markRunStart_();
    clearTab_(EVENTS, EVENT_HEADER);
  }
  catch (e) {
   ui.alert("ERROR - unhandled exception: " + e);
  }
  finally {
    markRunEnd_();
  }
  // }
}

function menuItem10_() {
  
  let ui = SpreadsheetApp.getUi();
  
  if (needToAbortRun_(ui) || !dependenciesGood()) {
    return;
  }
  
  
  let result = ui.alert('Please confirm the event "unveiling"', 
                        'This will overwrite what is currently in the Review tab with events that have their reference-id fields replaced with names for accounts, partners, opportunities and users. You will loose any updates currently in the Review tab!\n\nYou want to continue anyway?',
                        ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {  
    
    try {
      markRunStart_();
      unveil_se_events();
    }
    catch (e) {
      ui.alert("ERROR - unhandled exception: " + e);
    }
    finally {
      markRunEnd_();
    }
  }
}

function menuItem11_() {
  
  let ui = SpreadsheetApp.getUi();
  
  if (needToAbortRun_(ui)) {
    return;
  }
  
  
  let result = ui.alert('Please confirm', 
                        'You want to clear the Review tab? All updates to reviewed events will be lost!',
                        ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {
    
    try {
      markRunStart_();
      clearTab_(EVENTS_UNVEILED, REVIEW_HEADER);
      
      let spreadsheet = SpreadsheetApp.getActive().getSheetByName(EVENTS_UNVEILED);
      spreadsheet.getRange('2:1000').setBackground('#ffffff'); 
    }
    catch (e) {
      ui.alert("ERROR - unhandled exception: " + e);
    }
    finally {
      markRunEnd_();
    }
  }
}

function menuItem12_() {
  
  let ui = SpreadsheetApp.getUi();
  
  if (needToAbortRun_(ui)) {
    return;
  }
  
  try {
    markRunStart_();
    clearTab_(LOG_TAB);
  }
  catch (e) {
    ui.alert("ERROR - unhandled exception: " + e);
  }
  finally {
    markRunEnd_();
  }
  
}

function menuItem13_() {
  
  let ui = SpreadsheetApp.getUi();
  
  if (needToAbortRun_(ui)) {
    return;
  }
  
  
  let result = ui.alert('Please confirm', 
                        'You want to reconcile the Review tab with the Events tab? Records in the Events tab will be updated to match corresponding records in the Review tab that have been modified.\n\nContinue?',
                        ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {
    
    try {
      markRunStart_();
      reconcile_se_events();
    }
    catch (e) {
      ui.alert("ERROR - unhandled exception: " + e);
    }
    finally {
      markRunEnd_();
    }
  }
}

function menuItem14_() {
  
  let ui = SpreadsheetApp.getUi();
  
  if (needToAbortRun_(ui)) {
    return;
  }
  
  let result = ui.alert('Please confirm delete of EVERYTHING!', 
                        'Do you want to clear out all calendar and event data?\n\nThe upload zap must be OFF before you continue!\n\nIs Zapier off? Do you want to do this?',
                        ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {
    try {
      markRunStart_();
      clearTab_(UPLOAD_STAGE, UPLOAD_STAGE_HEADER);
      clearTab_(CALENDAR, CALENDAR_HEADER);
      clearTab_(EVENTS, EVENT_HEADER);
      clearTab_(EVENTS_UNVEILED, REVIEW_HEADER);
      clearTab_(LOG_TAB);
    }
    catch (e) {
      ui.alert("ERROR - unhandled exception: " + e);
    }
    finally {
      markRunEnd_();
    }
  }
}

function menuItem15_() {
  
  let ui = SpreadsheetApp.getUi();
  
  if (needToAbortRun_(ui)) {
    return;
  }
  
  let result = ui.alert('Please confirm configuration delete!', 'Do you want to clear out ALL configuration data (the Salesforce info in the blue tabs)?',
                        ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {
    try {
      markRunStart_();
      clearTab_(IN_REGION_CUSTOMERS, CUSTOMER_HEADER);
      clearTab_(ALL_CUSTOMERS, CUSTOMER_HEADER);
      clearTab_(LEADS, LEAD_HEADER);     
      clearTab_(PARTNERS, PARTNER_HEADER);
      clearTab_(STAFF);
      clearTab_(OPPORTUNITIES);
      clearTab_(HISTORY);
    }
    catch (e) {
      ui.alert("ERROR - unhandled exception: " + e);
    }
    finally {
      markRunEnd_();
    }
  }
}

function menuItem20_() {
  
  let ui = SpreadsheetApp.getUi();
  ui.alert("Import Calendars\n\nThis will import invites from the Google Calendar app into the 'Calendar' tab of this sheet for each user specified in the 'Calendars' block of the 'Run Settings' tab. " + 
  "It will clear whatever is currently in the 'Calendar' tab before importing a fresh set of invites. Basic stats will be output to the 'Log' tab.");

}

function menuItem22_() {
  
  let ui = SpreadsheetApp.getUi();
  ui.alert("Generate Events\n\nThis will generate records in the 'Events' tab for each invite in the 'Calendar' tab that includes a customer, partner or lead. " + 
           "It will clear whatever is currently in the 'Events' tab before processing invites. Run information will be output to the 'Log' tab, including any invites that couldn't be processed. " +
           "Please see the \"WARNING - Unable to find a customer, partner or lead for:\" log message in the manual.");
  
}

function menuItem23_() {
  
  let ui = SpreadsheetApp.getUi();
  ui.alert("Unveil Events\n\nThis takes each record in the 'Events' tab, replacing the Salesforce IDs in user, customer, partner, lead and opportunity reference fields with the corresponding names that people can understand, and writes those 'unveiled' records into the 'Review' tab. " +
           "It will clear whatever is currently in the 'Review' tab before performing the ID to name expansion. These unveiled records are easier to manually review prior to upload.");
  
}

function menuItem24_() {
  
  let ui = SpreadsheetApp.getUi();
  ui.alert("Process Events\n\nExecute the three pre-upload event processing stages in order:\n\n1. Import Calendars\n2. Generate Events\n3. Unveil Events.\n\nMay take a couple minutes to complete, so go get some coffee.");
  
}

function menuItem25_() {
  
  let ui = SpreadsheetApp.getUi();
  ui.alert("Upload Events\n\nThis takes each record in the 'Events' tab and appends it into the 'Uploads' tab. " +
           "From there a zap will upload the freshly appended records to Salesforce. Ensure that you review the events for correctness before uploading to Salesforce!");
  
}

function menuItem26_() {
  
  let ui = SpreadsheetApp.getUi();
  ui.alert("Reconcile Events\n\nThis takes each record in the 'Review' tab that has been modified during the review process and updates the corresponding record back in the 'Events' tab. " +
           "Note that the stage field may be overridden if Salesforce API validation would fail for the selected meeting type (review the Log for any overrides.)");
  
}

function menuItem27_() {
  
  let ui = SpreadsheetApp.getUi();
  ui.alert("Check Version\n\nUse this after pulling new code from Github. It will let you know if you are good to go, and update the version of the sheet to match the code. " +
           "If you are not good, it will tell you what you need to do, i.e. go get the latest copy of the spreadsheet and reconfigure from there."); 
}

function menuItem30_() {
  
  let ui = SpreadsheetApp.getUi();
   
  let sheet = SpreadsheetApp.getActive().getSheetByName(RUN_PARMS);
  let gasRange = sheet.getRange(1,1,1,1); // Cell A1
  let gasValues = gasRange.getValues();
  let gasVersionString = gasValues[0][0];
  let gasVersion = parseInt(gasVersionString.substring(1).replace(/\./g, ""));  // Assumes v#.#.# in cell A1
  if (GAS_VERSION == gasVersion) {
    let theMessage = "Sheet version matches the GAS (Google App Script). You are good to go.";
    ui.alert(theMessage);
    return;
  }  
  
  sheet = SpreadsheetApp.getActive().getSheetByName(CHOICES);
  let schemaRange = sheet.getRange(4,6,1,1); // Cell F4
  let schemaValues = schemaRange.getValues();
  let schemaVersionString = schemaValues[0][0];
  let schemaVersion = 0;
  if (schemaVersionString) {
    schemaVersion = parseInt(schemaVersionString.replace(/\./g, ""));
  }
  
  if (schemaVersion < MIN_SCHEMA_VERSION) {
    ui.alert("The schema of your sheet is too old! The GAS (Google App Script) you have pulled from Github won't work with it. " +
             "You need to start fresh.\n\nI'm sorry. Sometimes progress requires adding or changing some tabs or cells. Github only manages the code.\n\n" +
             "Please copy \"Activity Machine " + GAS_VERSION_STRING + "\" to you local Google drive, update the \"Upload Event\" zap to use it, and reload your config.");
  }
  else if (GAS_VERSION != gasVersion) {
    ui.alert("Sheet schema is compatible with the GAS (Google App Script). Updating the sheet version to " + "v" + GAS_VERSION_STRING);
    gasRange.setValue("v" + GAS_VERSION_STRING);
  }
  if (gasVersion < 100) {
  
    let m = "ZAP NOTICES\n\nBoth the \"Import Missing Accounts\" and \"Import Missing Leads\" Zaps have been DEPRECATED!\n(And there was much rejoicing!) " +
    "You can delete them!\n\nThey have been replaced by blue configuration tabs for ALL customers and ALL leads in Salesforce. " +
    "(Logic was added to handle massive numbers of records in a config tab.)";
    
    if (gasVersion < 18) {
    
      m+= "\n\nAlso note that the \"Upload SE Events\" zap was modified in version 0.1.8 to upload the meeting \"Notes URL\". " +
      "See the manual, addendum A in the Setup section, for the modification procedure.";
    
    }
    
    ui.alert(m);
  }
}


function onEdit(e) {
  
  let thisSheet = e.source.getActiveSheet();
  if (thisSheet.getName() == RUN_PARMS && e.range.rowStart == e.range.rowEnd) {
    handleRunParmsMenuEdit_(e.range, e.value);
  }
  else if (thisSheet.getName() == EVENTS_UNVEILED && !inReviewInitialization) {

    if (!e.value && e.range.rowStart == e.range.rowEnd) {
      // This means a range inside of a filtered review table was updated. The event only 
      // Passes in the first cell, even though many cells, not necessarily contigous, are updated.
      // In this case we can't track which rows, so we have to turn off the dirty row tracking and
      // fall back to processing every single row on a reconcile. 
      // There must be some way to determine the other cells, but Google is making it incredibly hard to figure out.
      SpreadsheetApp.getUi().alert("WARNING\n\nA filtered range on the review tab was updated. Color coding to indicate updated values MAY or MAY NOT be completely accurate.\n\nReconcile should work however.");
      PropertiesService.getScriptProperties().setProperty("reviewTouchEnabled", "false"); // turned out this didn't save much time, but ... was cool so leaving it in. ;)
    }
    else {
      updateDirtyRows(e.range.rowStart, e.range.rowEnd); // Filter on rows for follow-on reconcile (so we don't waste time copying every row from review back to events)
      handleReviewMenuEdit_(e.range, e.value); // Just sets the red indicator for cell change
    }
  }
}

function updateDirtyRows(start, end) {
  
  if (start > end) {
    // GAS has gone brain dead
    logOneCol("ERROR - Google range corrupted during review! A subsequent reconcile may not be complete.");
    return;
  }
  
  try {
    var lock = LockService.getScriptLock(); // In case two people are making updates to the review tab at the same time
    lock.waitLock(5000); // 5 sec
    let reviewRowWasTouchedArray = JSON.parse(PropertiesService.getScriptProperties().getProperty("reviewTouches"));
    for (let i = start; i < end + 1; i++) {
      reviewRowWasTouchedArray[i] = true;   
    } 
    PropertiesService.getScriptProperties().setProperty("reviewTouches", JSON.stringify(reviewRowWasTouchedArray));     
  }
  catch (err) {
    logOneCol("WARNING - Exeception during review update tracking. A subsequent reconcile may not be complete.");
  }
  finally {
    lock.releaseLock();
  }
}

/*
The event e ...
{ 
    String user, 
    SpreadSheet source, 
    Range range,
    Object value 
}
*/

function handleRunParmsMenuEdit_(cell, value) {
  
  // Columns an Rows hardcoded to Account Overrides name entries
  if (cell.columnStart != 7) return;
  var watchRows = [4, 5, 6, 7, 8, 9, 10, 11, 12, 15, 16, 17, 18, 19]; 
  if (watchRows.indexOf(cell.rowStart) == -1) return;
  
  if (!value) {
    cell.offset(0, 1).setValue("");
    return;
  }
  
  var tab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(IN_REGION_CUSTOMERS);
  var range = tab.getRange(2, 2, 500, 2); // 500 accounts max. First account name is on row 2, column 4. Id is at column 2.
  //SpreadsheetApp.getUi().alert("here we go");
  for (var i = 0; i < 500; i++) {
    if (range.offset(i, 2).getValue() == value) {
      let v = range.offset(i,0).getValue();
      cell.offset(0, 1).setValue(v); 
      // cell.offset(0, 1).setValue(v.substring(0, v.length - 3)); // Creating "short" id (cutting off 3 character post fix)
      break;
    }
  }
}

function handleReviewMenuEdit_(range, value) {

  if (range.rowStart > MAX_ROWS) {
    SpreadsheetApp.getUi().alert("We have exceeded " + MAX_ROWS + " events! Please reduce the number of calendars or the length of time being analyzed, or talk to Dave about bumping this limit up. Thank you.\n--The Activity Machine");
    return; 
  }  
 
  if (range.rowStart == range.rowEnd && range.columnStart == range.columnEnd) {
    // value only set if range is a single cell
    if (range.columnStart == REVIEW_EVENT_TYPE + 1) {
      handleEventTypeChange_(range, value);
    }
    else {
      let origValue = range.offset(0, getOriginalValueOffset_(range.columnStart)).getValue();
      colorCell_(range, value, origValue);
    }
  }
  else {
    for (let i = 0; i <= range.rowEnd - range.rowStart; i++) {
      for (let j = 0; j <= range.columnEnd - range.columnStart; j++) {
        let cell = range.offset(i,j,1,1);
        let value = cell.getValue();
        if (range.columnStart + j == REVIEW_EVENT_TYPE + 1) {
          handleEventTypeChange_(cell, value);
        }
        else {
          let origValue = cell.offset(0, getOriginalValueOffset_(range.columnStart + j)).getValue();
          colorCell_(cell, value, origValue);
        }
      }
    }
  }
}

function getOriginalValueOffset_(column) {

  let retVal = 0;
  switch (column) {
    case REVIEW_RELATED_TO + 1:
      retVal = (REVIEW_ORIG_RELATED_TO - REVIEW_RELATED_TO);
      break;
    case REVIEW_OP_STAGE + 1:
      retVal = (REVIEW_ORIG_OP_STAGE - REVIEW_OP_STAGE);
      break;
    case REVIEW_PRODUCT + 1:
      retVal = (REVIEW_ORIG_PRODUCT - REVIEW_PRODUCT);
      break;
    case REVIEW_DESC + 1:
      retVal = (REVIEW_ORIG_DESC - REVIEW_DESC);
      break;
    case REVIEW_MEETING_TYPE + 1:
      retVal = (REVIEW_ORIG_MEETING_TYPE - REVIEW_MEETING_TYPE);
      break;
    case REVIEW_REP_ATTENDED + 1:
      retVal = (REVIEW_ORIG_REP_ATTENDED - REVIEW_REP_ATTENDED);
      break;
    case REVIEW_LOGISTICS + 1:
      retVal = (REVIEW_ORIG_LOGISTICS - REVIEW_LOGISTICS);
      break;
    case REVIEW_PREP_TIME + 1:
      retVal = (REVIEW_ORIG_PREP_TIME - REVIEW_PREP_TIME);
      break;
    case REVIEW_QUALITY + 1:
      retVal = (REVIEW_ORIG_QUALITY - REVIEW_QUALITY);
      break;
    case REVIEW_LEAD + 1:
      retVal = (REVIEW_ORIG_LEAD - REVIEW_LEAD);
      break;
    case REVIEW_NOTES + 1:
      retVal = (REVIEW_ORIG_NOTES - REVIEW_NOTES);
      break;
    //case REVIEW_ACCOUNT_TYPE + 1:
    //  colorCell_(cell, value, REVIEW_ORIG_ACCOUNT_TYPE - REVIEW_ACCOUNT_TYPE);
    //  break;
    case REVIEW_PROCESS + 1:
      retVal = (REVIEW_ORIG_PROCESS - REVIEW_PROCESS);
      break;
    default:
  }
  return retVal;
}


function handleEventTypeChange_(cell, value) {
 
  let validationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHOICES);
  let validationRange = validationSheet.getRange(4, 4, 2); // Null choices for "Unknown"
  let validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
  let clearedValidationRule = validationRule;
  let validationOffset = REVIEW_RELATED_TO - REVIEW_EVENT_TYPE;
  let clearingOffset = REVIEW_LEAD - REVIEW_EVENT_TYPE;
  let originalValueOffset = REVIEW_ORIG_LEAD - REVIEW_LEAD + clearingOffset;
  
  switch (value) {
      
    case "Unknown":
      break;
    case "Opportunity":
      validationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHOICE_OP);
      validationRange = validationSheet.getRange(2, CHOICE_OP_NAME+1, validationSheet.getLastRow()); // Skip 1 row header
      validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
      break;
    case "Partner":
      validationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHOICE_PARTNER);
      validationRange = validationSheet.getRange(2, CHOICE_PARTNER_NAME+1, validationSheet.getLastRow()); // Skip 1 row header
      validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
      break;
    case "Customer":
      validationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHOICE_ACCOUNT);
      validationRange = validationSheet.getRange(2, CHOICE_ACCOUNT_NAME+1, validationSheet.getLastRow()); // Skip 1 row header
      validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
      break;
    case "Lead":
      validationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHOICE_LEAD);
      validationRange = validationSheet.getRange(2, CHOICE_LEAD_EMAIL+1, validationSheet.getLastRow()); // Skip 1 row header
      validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
      validationOffset = REVIEW_LEAD - REVIEW_EVENT_TYPE;
      clearingOffset = REVIEW_RELATED_TO - REVIEW_EVENT_TYPE;
      originalValueOffset = REVIEW_ORIG_RELATED_TO - REVIEW_RELATED_TO + clearingOffset; 
      break;
    default:
  }
   
  cell.offset(0, clearingOffset).setValue("");
  cell.offset(0, clearingOffset).clearDataValidations(); 
  let origValue = cell.offset(0, originalValueOffset).getValue();
  colorCell_(cell.offset(0, clearingOffset), "", origValue);
  cell.offset(0, validationOffset).setDataValidation(validationRule);
  origValue = cell.offset(0, REVIEW_ORIG_EVENT_TYPE - REVIEW_EVENT_TYPE).getValue();
  colorCell_(cell, value, origValue);
  
}

function initValidation_(cell, type) {
  
  let validationSheet = null;
  let validationRange = null;
  let validationRule = null;
  
  switch (type) {
      
    case "Unknown":
      return;
      break;
    case "Opportunity":
      validationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHOICE_OP);
      validationRange = validationSheet.getRange(2, CHOICE_OP_NAME+1, validationSheet.getLastRow()); 
      validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
      break;
    case "Partner":
      validationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHOICE_PARTNER);
      validationRange = validationSheet.getRange(2, CHOICE_PARTNER_NAME+1, validationSheet.getLastRow()); 
      validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
      break;
    case "Customer":
      validationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHOICE_ACCOUNT);
      validationRange = validationSheet.getRange(1, CHOICE_ACCOUNT_NAME+1, validationSheet.getLastRow()); 
      validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
      break;
    case "Lead":
      validationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LEADS);
      validationRange = validationSheet.getRange(2, LEAD_NAME+1, validationSheet.getLastRow()); 
      validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
      break;
    default:
      Logger.log("ERROR: bogus value in Review tab's Event Type field: " + value);
      return;
  }
  
  cell.setDataValidation(validationRule);
}



function colorCell_(cell, value, originalValue) {
  if (typeof value === 'undefined') value = "";
  //let originalValue = cell.offset(0, originalValueOffset).getValue();
  if (originalValue != value) {
    cell.setBackground("tomato");
  }
  else {
    cell.setBackground(null);
  }
}

