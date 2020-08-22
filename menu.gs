function markRunStart_() {
  PropertiesService.getScriptProperties().setProperty('whoIsRunning', Session.getActiveUser().getEmail());
}

function markRunEnd_() {
  PropertiesService.getScriptProperties().deleteProperty('whoIsRunning');
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
    Logger.log("ERROR: import_calendar threw an unhandled exception!");
  }
  finally {
    markRunEnd_();
  }
  // }
}

function menuItem2_() {
  
  let ui = SpreadsheetApp.getUi();
  
  if (needToAbortRun_(ui)) {
    return;
  }
  
  
  let result = ui.alert('Please confirm', 
                        'Are you sure you want to transform invites into events? This will clear the current set of Events and Expaned/Reviewed events.\n\nContinue?',
                        ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {  
    try {
      markRunStart_();
      build_se_events();
    }
    catch (e) {
      Logger.log("ERROR: build_se_events threw an unhandled exception!");
    }
    finally {
      markRunEnd_();
    }
  }
}

function menuItem3_() {
  
  let ui = SpreadsheetApp.getUi();
  
  if (needToAbortRun_(ui)) {
    return;
  }
  
  let result = ui.alert('Please confirm', 
                        'Are you sure you want to import invites and transform into events? This will clear the current set of invites and events, replace them with a fresh import, and fresh set of events from that import.\n\nContinue?',
                        ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {
    try {
      markRunStart_();
      import_calendars();
      build_se_events();
    }
    catch (e) {
      Logger.log("ERROR: import_calendar or build_se_events threw an unhandled exception!");
    }
    finally {
      markRunEnd_();
    }
  }
}

function menuItem4_() {
  
  let ui = SpreadsheetApp.getUi();
  if (needToAbortRun_(ui)) {
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
      Logger.log("ERROR: stage_events_to_upload_tab threw an unhandled exception!");
    }
    finally {
      markRunEnd_();
    }
  }
}

function menuItem5_() {
  
  let ui = SpreadsheetApp.getUi();
  
  if (needToAbortRun_(ui)) {
    return;
  }
  
  let result = ui.alert('Please confirm', 
                        'Are you sure you want to clear the "Missing Tabs"? The Zapier "Import Missing Accounts" and "Import Missing Leads" zaps must be OFF!\n\nAre they off?',
                        ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) { 
    try {
      markRunStart_();
      clearTab_(MISSING_CUSTOMERS, CUSTOMER_HEADER);
      clearTab_(MISSING_LEADS, LEAD_HEADER);
      clearTab_(MISSING_DOMAINS, [['Email']]); 
      clearTab_(MISSING_EMAILS, [['Email']]); 
    }
    catch (e) {
      Logger.log("ERROR: clearTab_ threw an unhandled exception!");
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
      Logger.log("ERROR: clearTab_ threw an unhandled exception!");
    }
    finally {
      markRunEnd_();
    }
  }
}

function menuItem7_() {
  
  let ui = SpreadsheetApp.getUi();
  
  if (needToAbortRun_(ui)) {
    return;
  }
  
  try {
    import_missing_accounts();
  }
  catch (e) {
    Logger.log("ERROR: import_missing_accounts threw an exception!: " + e);
  }
  finally {
    markRunEnd_();
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
    Logger.log("ERROR: clearTab_ threw an unhandled exception!");
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
    Logger.log("ERROR: clearTab_ threw an unhandled exception: " + e);
  }
  finally {
    markRunEnd_();
  }
  // }
}

function menuItem10_() {
  
  let ui = SpreadsheetApp.getUi();
  
  if (needToAbortRun_(ui)) {
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
      Logger.log("ERROR: unveil_se_events threw an unhandled exception: " + e);
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
      Logger.log("ERROR: clearTab_ threw an unhandled exception: " + e);
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
    Logger.log("ERROR: clearTab_ threw an unhandled exception: " + e);
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
      Logger.log("ERROR: reconcile_se_events threw an unhandled exception: " + e);
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

function menuItem21_() {
  
  let ui = SpreadsheetApp.getUi();
  ui.alert("Import Missing Accounts\n\nThis is a dynamic configuration update method. It will identify domains not represented in the current 'In Region Customers' tab along with the email addresses of potential leads for those missing domains, " +
  "appending that information into the 'Missing Domains' and 'Missing Emails' tabs. From there, two zaps will search for those recently added 'missing' customers and leads, appending what it finds, " +
  "if anything, into the 'Missing Customers' and 'Missing Leads' tabs. This should be run after a fresh Calendar Import.");

}

function menuItem22_() {
  
  let ui = SpreadsheetApp.getUi();
  ui.alert("Generate Events\n\nThis will generate records in the 'Events' tab for each invite in the 'Calendar' tab that includes a customer, partner or lead. " + 
           "It will clear whatever is currently in the 'Events' tab before processing invites. Run information will be output to the 'Log' tab, including which 'externally facing' invites could not be matched with a customer, partner or lead.");
  
}

function menuItem23_() {
  
  let ui = SpreadsheetApp.getUi();
  ui.alert("Unveil Events\n\nThis takes each record in the 'Events' tab, replacing the Salesforce IDs in user, customer, partner, lead and opportunity reference fields with the corresponding names that people can understand, and writes those 'unveiled' records into the 'Review' tab. " +
           "It will clear whatever is currently in the 'Review' tab before performing the ID to name expansion. These unveiled records are easier to manually review prior to upload.");
  
}

function menuItem24_() {
  
  let ui = SpreadsheetApp.getUi();
  ui.alert("Import Calendars & Generate Events\n\nThis is a convenience method to execute Import Calendars then Generate Events in sequence. " +
           "Note that Import Missing Accounts will not be executed, so only execute this if you know there are no missing accounts for the current calendar (you have likely searched for missing accounts recently.)");
  
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

function onEdit(e) {
   
  // Only do this is a single cell is changed
  if (e.range.rowStart != e.range.rowEnd || e.range.columnStart != e.range.columnEnd) {
    return;
  }
  
  let thisSheet = e.source.getActiveSheet();
  if (thisSheet.getName() == RUN_PARMS) {
    handleRunParmsMenuEdit_(e.range, e.value);
  }
  else if (thisSheet.getName() == EVENTS_UNVEILED) {
    handleReviewMenuEdit_(e.range, e.value);
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
      cell.offset(0, 1).setValue(v.substring(0, v.length - 3)); // Creating "short" id (cutting off 3 character post fix)
      break;
    }
  }
}

function handleReviewMenuEdit_(cell, value) {

  if (cell.rowStart > MAX_ROWS) {
    SpreadsheetApp.getUi().alert("We have exceeded " + MAX_ROWS + " events! Please reduce the number of calendars or the length of time being analyzed, or talk to Dave about bumping this limit up. Thank you.\n--The Activity Machine");
    return; 
  }  
  if (!value) {
    Logger.log("ERROR: data validation failing on Review tab's Event Type field!");
    return;
  }
  switch (cell.columnStart) {
    case REVIEW_EVENT_TYPE + 1:
      handleEventTypeChange_(cell, value);
      break;
    case REVIEW_RELATED_TO + 1:
      handleValueSelection_(cell, value, REVIEW_ORIG_RELATED_TO - REVIEW_RELATED_TO);
      break;
    case REVIEW_OP_STAGE + 1:
      handleValueSelection_(cell, value, REVIEW_ORIG_OP_STAGE - REVIEW_OP_STAGE);
      break;
    case REVIEW_PRODUCT + 1:
      handleValueSelection_(cell, value, REVIEW_ORIG_PRODUCT - REVIEW_PRODUCT);
      break;
    case REVIEW_DESC + 1:
      handleValueSelection_(cell, value, REVIEW_ORIG_DESC - REVIEW_DESC);
      break;
    case REVIEW_MEETING_TYPE + 1:
      handleValueSelection_(cell, value, REVIEW_ORIG_MEETING_TYPE - REVIEW_MEETING_TYPE);
      break;
    case REVIEW_REP_ATTENDED + 1:
      handleValueSelection_(cell, value, REVIEW_ORIG_REP_ATTENDED - REVIEW_REP_ATTENDED);
      break;
    case REVIEW_LOGISTICS + 1:
      handleValueSelection_(cell, value, REVIEW_ORIG_LOGISTICS - REVIEW_LOGISTICS);
      break;
    case REVIEW_PREP_TIME + 1:
      handleValueSelection_(cell, value, REVIEW_ORIG_PREP_TIME - REVIEW_PREP_TIME);
      break;
    case REVIEW_QUALITY + 1:
      handleValueSelection_(cell, value, REVIEW_ORIG_QUALITY - REVIEW_QUALITY);
      break;
    case REVIEW_LEAD + 1:
      handleValueSelection_(cell, value, REVIEW_ORIG_LEAD - REVIEW_LEAD);
      break;
    default:
  }
}


function handleEventTypeChange_(cell, value) {
 
  let validationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHOICES);
  let validationRange = validationSheet.getRange(4, 4, 2); // Null choices for "Unknown"
  let validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
  let clearedValidationRule = validationRule;
  let validationOffset = REVIEW_RELATED_TO - REVIEW_EVENT_TYPE;
  let clearingOffset = REVIEW_LEAD - REVIEW_EVENT_TYPE;
  let originalValueOffset = REVIEW_ORIG_LEAD - REVIEW_LEAD;
  
  
  switch (value) {
      
    case "Unknown":
      break;
    case "Opportunity":
      validationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OPPORTUNITIES);
      validationRange = validationSheet.getRange(2, OP_NAME+1, validationSheet.getLastRow()); // Skip 1 row header
      validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
      break;
    case "Partner":
      validationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTNERS);
      validationRange = validationSheet.getRange(2, PARTNER_NAME+1, validationSheet.getLastRow()); // Skip 1 row header
      validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
      break;
    case "Customer":
      validationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(IN_REGION_CUSTOMERS);
      validationRange = validationSheet.getRange(2, CUSTOMER_NAME+1, validationSheet.getLastRow()); // Skip 1 row header
      validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
      break;
    case "Lead":
      validationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MISSING_LEADS);
      validationRange = validationSheet.getRange(2, LEAD_NAME+1, validationSheet.getLastRow()); // Skip 1 row header
      validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
      validationOffset = REVIEW_LEAD - REVIEW_EVENT_TYPE;
      clearingOffset = REVIEW_RELATED_TO - REVIEW_EVENT_TYPE;
      originalValueOffset = REVIEW_ORIG_RELATED_TO - REVIEW_RELATED_TO;
      break;
    default:
      Logger.log("ERROR: bogus value in Review tab's Event Type field: " + value);
  }
  
  
  cell.offset(0, clearingOffset).setValue("");
  cell.offset(0, clearingOffset).clearDataValidations(); 
  handleValueSelection_(cell.offset(0, clearingOffset), "", originalValueOffset);
  cell.offset(0, validationOffset).setDataValidation(validationRule);
  handleValueSelection_(cell, value, REVIEW_ORIG_EVENT_TYPE - REVIEW_EVENT_TYPE);
  
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
      validationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OPPORTUNITIES);
      validationRange = validationSheet.getRange(2, OP_NAME+1, validationSheet.getLastRow()); // Skip 1 row header
      validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
      break;
    case "Partner":
      validationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTNERS);
      validationRange = validationSheet.getRange(2, PARTNER_NAME+1, validationSheet.getLastRow()); // Skip 1 row header
      validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
      break;
    case "Customer":
      validationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(IN_REGION_CUSTOMERS);
      validationRange = validationSheet.getRange(2, CUSTOMER_NAME+1, validationSheet.getLastRow()); // Skip 1 row header
      validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
      break;
    case "Lead":
      validationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MISSING_LEADS);
      validationRange = validationSheet.getRange(2, LEAD_NAME+1, validationSheet.getLastRow()); // Skip 1 row header
      validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
      break;
    default:
      Logger.log("ERROR: bogus value in Review tab's Event Type field: " + value);
      return;
  }
  
  cell.setDataValidation(validationRule);
}

function handleValueSelection_(cell, value, originalValueOffset) {
  let originalValue = cell.offset(0, originalValueOffset).getValue();
  if (value != originalValue) {
    let reviewRowWasTouchedArray = PropertiesService.getScriptProperties().getProperty("reviewTouches");
    reviewRowWasTouchedArray[cell.rowStart] = true;   
    PropertiesService.getScriptProperties().setProperty("reviewTouches", reviewRowWasTouchedArray); 
    cell.setBackground("tomato");
  }
  else {
    cell.setBackground(null);
  }
}

