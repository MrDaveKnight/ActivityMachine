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
  
  /*
  let result = ui.alert('Please confirm', 
  'Are you sure you want to transform invites into events? This will clear the current set of Events.\n\nContinue?',
  ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {  
  */
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
  //}
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
      clearTab_(UPLOAD_STAGE, EVENT_HEADER);
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
    Logger.log("ERROR: clearTab_ threw an unhandled exception!");
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
  
  /*
  let result = ui.alert('Please confirm event field "expansion"', 
  'You want to transform the account, partner, opportunity and user ids in Events into names and save in the Expanded tab?',
  ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {  
  */
  try {
    markRunStart_();
    expand_se_events();
  }
  catch (e) {
    Logger.log("ERROR: expand_se_events threw an unhandled exception!");
  }
  finally {
    markRunEnd_();
  }
  // }
}

function menuItem11_() {
  
  let ui = SpreadsheetApp.getUi();
  
  if (needToAbortRun_(ui)) {
    return;
  }
  
  /*
  let result = ui.alert('Please confirm', 
  'You want to clear the Expanded event tab?',
  ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {
  */
  try {
    markRunStart_();
    clearTab_(EVENTS_EXPANDED, EVENT_HEADER);
  }
  catch (e) {
    Logger.log("ERROR: clearTab_ threw an unhandled exception!");
  }
  finally {
    markRunEnd_();
  }
  //}
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
    Logger.log("ERROR: clearTab_ threw an unhandled exception!");
  }
  finally {
    markRunEnd_();
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
  ui.alert("Expand Events\n\nThis takes each record in the 'Events' tab, replacing Salesforce IDs with corresponding names, and writes those 'expanded' records into the 'Expanded' tab. " +
  "It will clear whatever is currently in the 'Expanded' tab before performing the ID to name expansion. These expanded records are easier to manually review prior to upload.");

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

function onEdit(e) {
  var thisSheet = e.source.getActiveSheet();
  
  
  if (thisSheet.getName() != RUN_PARMS) {
    return;
  }

  // Hardcoded to Account Overrides name entries
  var watchColumns = [7];
  if (watchColumns.indexOf(e.range.columnStart) == -1) return;
  var watchRows = [4, 5, 6, 7, 8, 9, 10, 11, 12, 15, 16, 17, 18, 19]; 
  if (watchRows.indexOf(e.range.rowStart) == -1) return;

  if (!e.value) {
    e.range.offset(0, 1).setValue("");
    return;
  }

  var tab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(IN_REGION_CUSTOMERS);
  var range = tab.getRange(2, 2, 500, 2); // 500 accounts max. First account name is on row 2, column 4. Id is at column 2.
  //SpreadsheetApp.getUi().alert("here we go");
  for (var i = 0; i < 500; i++) {
    if (range.offset(i, 2).getValue() == e.value) {
      let v = range.offset(i,0).getValue();
      e.range.offset(0, 1).setValue(v.substring(0, v.length - 3)); // Creating "short" id (cutting off 3 character post fix)
      break;
    }
  }
}

