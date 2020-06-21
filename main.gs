// Calendar Import
//
// This is a tool to automatically create SE Activities (Field Events in Salesforce) from calendar invites
// that include customers or partners. Internal meetings are not included. All customer/partner facing
// non-recurring invites are included. All customer/partner facing recurring invites that have been "accepted" 
// (SE has indicated they will be going) are also included.
// 
// The tool identifies the account(s) associated with each invite, along with the opportunity if any, based on the 
// attendiees and the products being discussed. The meeting type is derived from keywords in the invite
// such as "demo" or "pov". Reference the lookForMeetingType_() method for more detail.
//
// This script provides logic to -
// 1. Import google calendar invites into a google sheet
// 2. Generate SE Activities from those invites, saving them in a "staging" tab in the google sheet
// 3. Move the staged SE Activities into an "Upload" tab on the google sheet
//
// A Zapier zap is configured to watch for SE Activities that appear in the Upload tab, and push them up 
// into Salesforce
// 
// ** Remember to turn off the Zap before you delete records from the Upload tab. **
//
// Information from Salesforce needed by the invite to event transform logic is stored in other google sheet tabs.
//
// Written by Dave Knight, Rez Dogs 2020, knight@hashicorp.com

// Clear meeting type default
const MEETINGS_HAVE_DEFAULT=true;

// Information for accounts, staff and opportunities is imported directly from
// Salesforce and saved in spreadsheet tabs.
// Calendar invite information has been imported directly from google calendar into a tab.
// The indexes below identify the column locations of the imported data.

const RUN_PARMS = "Run Settings"

const MISSING_CUSTOMERS = "Missing Customers (Zap)";
const MISSING_EMAIL_DOMAIN = 0;

const IN_REGION_CUSTOMERS = "In Region Customers"; // Our Region's Customers or Prospects - Accounts for short. They are accounts in Salesforce
const EXTERNAL_CUSTOMERS = "External Customers (Zap)"; // Dynamic account imports for each missing customer
const CUSTOMER_COLUMNS = 7;
const CUSTOMER_ID = 1;
const CUSTOMER_NAME = 3;
const CUSTOMER_EMAIL_DOMAIN = 5;
const CUSTOMER_ALT_EMAIL_DOMAINS = 6;
const CUSTOMER_HEADER = [['Account Owner', '18-Digit Account ID', 'Solutions Engineer', 'Account Name', 'Owner Region', 'Email Domain', 'Alt Email Domains']];		

const PARTNERS = "Partners"; // Also an account in Salesforce
const PARTNER_COLUMNS = 4;
const PARTNER_ID = 1;
const PARTNER_NAME =2;
const PARTNER_EMAIL_DOMAIN = 3;
const PARTNER_ALT_EMAIL_DOMAINS = 4;
const PARTNER_HEADER = [['Account Owner', '18-Digit Account ID', 'Account Name', 'Email Domain']];			// FIXME Alt email															

const STAFF = "Staff"
const STAFF_COLUMNS = 5;
const STAFF_ID =3;
const STAFF_NAME = 0;
const STAFF_EMAIL = 1;
const STAFF_ROLE = 4;
const STAFF_ROLE_REP = "EAM 1"

const OPPORTUNITIES = "Opportunities";
const OP_COLUMNS = 11;
const OP_ID = 0;
const OP_OWNER = 3;
const OP_SE_PRIMARY = 1;
const OP_SE_SECONDARY = 2;
const OP_NAME = 7;
const OP_PRIMARY_PRODUCT = 9; // Appears to always be set
//const OP_LEAD_PRODUCT = 8;
const OP_CLOSE_DATE = 10;
const OP_ACCOUNT_ID = 11;
const OP_STAGE = 5;
const OP_TYPE = 6;

const HISTORY = "Op History";
const HISTORY_COLUMNS = 3;
const HISTORY_OP_ID = 0;
const HISTORY_STAGE_VALUE = 1;
const HISTORY_STAGE_DATE = 2;

const CALENDAR = "Calendar";
const CALENDAR_COLUMNS = 11;
const ASSIGNED_TO = 0;
const SUBJECT = 1;
const LOCATION = 2;
const START = 3;
const END = 4;
const ASSIGNEE_STATUS = 5;
const CREATED_BY = 6;
const IS_ALL_DAY = 7;
const IS_RECURRING = 8;
const ATTENDEE_STR = 9;
const DESCRIPTION = 10;
const CALENDAR_HEADER = [["Calendar Address", "Summary/Title", "Location", "Start", "End", "MyStatus", "Created By", "All Day Event", "Recurring Event", "Attendees", "Description"]] 

const UPLOAD_STAGE = "Upload (Zap)"
const EVENTS = "Events"
const EVENT_COLUMNS = 12;
const EVENT_ASSIGNED_TO = 0;
const EVENT_OP_STAGE = 1;
const EVENT_MEETING_TYPE = 2;
const EVENT_RELATED_TO = 3; // Opportunity or Account ID
const EVENT_SUBJECT = 4;
const EVENT_START = 5;
const EVENT_END = 6;
const EVENT_REP_ATTENDED = 7;
const EVENT_PRODUCT = 8;
const EVENT_DESC = 9;
const EVENT_LOGISTICS = 10;
const EVENT_PREP_TIME = 11;
const EVENT_HEADER = [["Assigned To", "Opportunity Stage", "Meeting Type", "Related To", "Subject", "Start", "End", "Rep Attended", "Primary Product", "Description", "Logistics", "Prep"]]

var emailToCustomerMap = {};
var emailToPartnerMap = {};
var staffEmailToIdMap = {};
var staffEmailToRoleMap = {};
var staffNameToEmailMap = {};
var opByCustomerAndProduct = {};
var numberOfOpsByCustomer = {};
var primaryOpByCustomer = {};
var productByOp = {};
var stageMilestonesByOp = {};
var accountTeamByOp = {};

const INTERNAL_CUSTOMER_TYPE = 1;
const EXTERNAL_CUSTOMER_TYPE = 2;
const PARTNER_TYPE = 3;
var accountType = {}; // Index by account ID

function printIt_() {
  let keys = PropertiesService.getScriptProperties().getKeys();
  for (var i = 0; i < keys.length; i++) {
    Logger.log(keys[i] + ":" + PropertiesService.getScriptProperties().getProperty(keys[i]));
  }
}

var missingAccounts = {}; // For tracking missing account data in lookForAccounts_()

// Product codes for table index keys
const TERRAFORM = "T";
const VAULT = "V";
const CONSUL = "C";
const NOMAD = "N";

// Debug tracing
const IS_TRACE_ACCOUNT_ON = false;
const TRACE_ACCOUNT_ID = "0011C00001rtkI3"
const TRACE_ACCOUNT_ID_LONG = "0011C00001rtkI3QAI"

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Activity')
      .addItem('Import Calendars', 'menuItem1_')
      .addSeparator()
      .addItem('Import Missing Accounts (& wait for zap)', 'menuItem7_')
      .addSeparator()
      .addItem('Generate Events', 'menuItem2_')
      .addSeparator()
      .addItem('Expand Events', 'menuItem10_')
      .addSeparator()
      .addItem('Import Calendars & Generate Events', 'menuItem3_')
      .addSeparator()
      .addItem('Upload Events (& wait for zap)', 'menuItem4_')
      .addSeparator()
      .addSubMenu(ui.createMenu('Clear')
          .addItem('Calendar Tab', 'menuItem8_')
          .addSeparator()
          .addItem('Event Tab', 'menuItem9_')
          .addSeparator()
          .addItem('External/Missing Customers Tabs (Turn off Zapper!)', 'menuItem5_')
          .addSeparator()
          .addItem('Upload Tab (Turn off Zapper!)', 'menuItem6_'))
      .addToUi();
}

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
  
  let result = ui.alert('Please confirm', 
                        'Are you sure you want to import calendar invites? This will clear the current set in the Calendar tab, and replace with a fresh import.\n\nContinue?',
                        ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {
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
  }
}

function menuItem2_() {
  
  let ui = SpreadsheetApp.getUi();
  
  if (needToAbortRun_(ui)) {
    return;
  }
  
  let result = ui.alert('Please confirm', 
                        'Are you sure you want to transform invites into events? This will clear the current set of Events.\n\nContinue?',
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
                        'Are you sure you want to clear the External and Missing Customers Tabs? The Zapier "Import Missing Accounts" zap must be OFF!\n\nIs it off?',
                        ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) { 
    try {
      markRunStart_();
      clearTab_(EXTERNAL_CUSTOMERS, CUSTOMER_HEADER);
      clearTab_(MISSING_CUSTOMERS, [['Email']]); 
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
  
  let result = ui.alert('Please confirm', 
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
    import_external_accounts();
  }
  catch (e) {
    Logger.log("ERROR: import_external_accounts threw an exception!: " + e);
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
  
  let result = ui.alert('Please confirm', 
                        'You want to clear the Calendar import tab?',
                        ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {
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
  }
}

function menuItem9_() {
  
  let ui = SpreadsheetApp.getUi();
  
  if (needToAbortRun_(ui)) {
    return;
  }
  
  let result = ui.alert('Please confirm', 
                        'You want to clear the Event staging tab?',
                        ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {
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
  }
}

function menuItem10_() {
  
  let ui = SpreadsheetApp.getUi();
  
  if (needToAbortRun_(ui)) {
    return;
  }
  
  let result = ui.alert('Please confirm event field "expansion"', 
                        'You want to transform the account, partner, opportunity and user ids in Events into names and save in the Expanded tab?',
                        ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {  
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
  }
}

//DAK
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


function import_calendars() {

  clearTab_(CALENDAR, CALENDAR_HEADER);
  
  let sheet = SpreadsheetApp.getActive().getSheetByName(CALENDAR);
  let importCursor = {sheet : sheet, row : 2}; // Drop row below the header we kept during the clear
  let parms = SpreadsheetApp.getActive().getSheetByName(RUN_PARMS);
  
  let dateRange = parms.getRange(3,2,2,1); // Hardcoded to format in RUN_PARMS
  let dates = dateRange.getValues();
  
  let calendarRange = parms.getRange(3,5,25,1); // Hardcoded to format in RUN_PARMS
  let calendars = calendarRange.getValues();
  
  for (var row in calendars) {
    if (calendars[row][0]) {      
      importRow = import_gcal_(calendars[row][0], dates[0][0], dates[1][0], importCursor);
    }
  }
}

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


function import_gcal_(calName, startDate, endDate, cursor){
  
  //
  // Export Google Calendar Events to a Google Spreadsheet
  //
  // This code retrieves events between 2 dates for the specified calendar.
  // It logs the results in the Calendar tab of the current spreadsheet starting at cell A2 ]
  //
  
  var cal = CalendarApp.getCalendarById(calName);
  
  if (!cal) {
    Logger.log("ERROR: not subscribed to " + calName);
    return;
  }
  
  // Optional variations on getEvents
  // var events = cal.getEvents(new Date("January 3, 2014 00:00:00 CST"), new Date("January 14, 2014 23:59:59 CST"));
  // var events = cal.getEvents(new Date("January 3, 2014 00:00:00 CST"), new Date("January 14, 2014 23:59:59 CST"), {search: 'word1'});
  // 
  // Explanation of how the search section works (as it is NOT quite like most things Google) as part of the getEvents function:
  //    {search: 'word1'}              Search for events with word1
  //    {search: '-word1'}             Search for events without word1
  //    {search: 'word1 word2'}        Search for events with word2 ONLY
  //    {search: 'word1-word2'}        Search for events with ????
  //    {search: 'word1 -word2'}       Search for events without word2
  //    {search: 'word1+word2'}        Search for events with word1 AND word2
  //    {search: 'word1+-word2'}       Search for events with word1 AND without word2
  //
  //var events = cal.getEvents(new Date(startDate + " 00:00:00 CST"), new Date(endDate + " 23:59:59 CST"), {search: '-Hangs'});
  var events = cal.getEvents(new Date(startDate), new Date(endDate), {search: '-Hangs'});
  
  var cc = 0;  // Customer count (individuals)
  var cns = {}; // Customer names for a subsequent count (company count)
  
  // Loop through all calendar events found and write them out to the Calendar tab
  for (var i=0;i<events.length;i++) {
  
    // Collect some quick stats
    var isCustEvent = false;
    var attendees=events[i].getGuestList();
    var attendeeList = "";
    if (attendees.length > 0) {
      attendeeList = attendees[0].getEmail();
      if (attendeeList.indexOf("@hashicorp.com") == -1) {
        isCustEvent = true;
        cns[attendees[0].getEmail().substring(attendees[0].getEmail().indexOf("@"))] = true;
        cc++;
      }
    }
    for (var j=1; j < attendees.length; j++) {
      attendeeList = attendeeList + "," + attendees[j].getEmail();
      if (!isCustEvent && attendees[j].getEmail().indexOf("@hashicorp.com") == -1) {
        isCustEvent = true;
        cns[attendees[j].getEmail().substring(attendees[j].getEmail().indexOf("@"))] = true;
        cc++;
      }
    }
    
    /*
    No tags show up for Notes
    let tags = events[i].getAllTagKeys();
    let tagNames = "";
    if (tags && tags.length > 0) {
      tagNames = tags[0];
      for (let x = 1; x < tags.length; x++) {
        tagNames += ", " + tags[x];
      }
    }
    */
    
    // Matching the "header=" entry above, this is the detailed row entry "details=", and must match the number of entries of the GetRange entry below
    let details=[[calName, events[i].getTitle(), events[i].getLocation(), events[i].getStartTime(), events[i].getEndTime(), events[i].getMyStatus(), events[i].getCreators(), events[i].isAllDayEvent(), events[i].isRecurringEvent(), attendeeList, events[i].getDescription()]];
    let range=cursor.sheet.getRange(cursor.row,1,1,11); // 11 must match detaills
    range.setValues(details);
    cursor.row++;
  }
  
  sheet = SpreadsheetApp.getActive().getSheetByName("Log");
  var lastRow = sheet.getLastRow();
  var log = sheet.getRange(lastRow+1,1);
  log.offset(0, 0).setValue(calName + " import stats");
  log.offset(1, 0).setValue("Customer/Partner Meetings:");
  log.offset(1, 1).setValue(cc);
  
  var size = 0, key;
  for (key in cns) {
    if (cns.hasOwnProperty(key)) {
      size++;
    }
  }
  
  log.offset(2, 0).setValue("Companies Present:");
  log.offset(2, 1).setValue(size);
}


//function onOpen() {
//  Browser.msgBox('App Instructions - Please Read This Message', '1) Click Tools then Script Editor\\n2) Read/update the code with your desired values.\\n3) Then when ready click Run export_gcal_to_gsheet from the script editor.', Browser.Buttons.OK);
//}

function lookForAccounts_(attendees, customerMap, partnerMap) {

  // attendees: array of attendee emails
  // Returns the following ...
  let rv = {};
  rv.customers = {}; // These are 18-Digit IDs! May have to prune off 3 digit postfix when used for an op "key"
  rv.partners = {}; // These are 18-Digit IDs! Same warning
  rv.others = {}; // These are just email domains
  rv.stats = {};
  rv.stats.customers = 0; // Number of customers in attendence
  rv.stats.customer_accounts = 0; // Number of unique customer accounts represented
  rv.stats.partners = 0; // Number of partners in attendence
  rv.stats.hashi = 0; // Number of hashicorp attendees
  rv.stats.other = 0; // Number of unidentified attendees

  
  for (var j = 0; j < attendees.length; j++) {
    var domain = attendees[j].substring(attendees[j].indexOf("@")+1).trim();    
    var accountId = "";
    var type = "hashi";
    if (domain == "gmail.com") continue; // Thanks to Fernando for putting gmail as the domain for a prospect called "Freelance"
    
    if (customerMap[domain]) {
      type = "customers";
      accountId = customerMap[domain];
    }
    else if (partnerMap[domain] && domain != "hashicorp.com") {
      type = "partners";
      accountId = partnerMap[domain];
    }
    else if (domain != "hashicorp.com") {
      type = "others";
      accountId = domain;
    }
    
    if ("hashi" == type) {
      rv.stats.hashi++;
    }
    else {
      rv.stats[type]++;
      if (!rv[type][accountId]) {
        rv[type][accountId] = 1;
        if (type == "customers") {
          rv.stats.customer_accounts++;
        }
      }
      else {
        rv[type][accountId]++;
      }
    }
  }
  
  return rv;
}



function lookForProducts_(text) {
  
  let returnValue = {
    hasTerraform : false,
    hasVault : false,
    hasConsul : false,
    hasNomad : false
  }
  
  if (!text) return returnValue;
  var x = text.toLowerCase();
  
  returnValue.hasTerraform = x.indexOf("terraform") != -1;
  if (!returnValue.hasTerraform) {
    let regex = RegExp("p?tf[ce]");
    returnValue.hasTerraform = regex.test(x);
  }
  if (!returnValue.hasTerraform) {
    returnValue.hasTerraform = x.indexOf("tf cloud") != -1;
  }
  returnValue.hasVault = x.indexOf("vault") != -1;
  returnValue.hasConsul = x.indexOf("consul") != -1;
  returnValue.hasNomad = x.indexOf("nomad") != -1;
  
  return returnValue;
  
}

function lookForMeetingType_(stage, text) {
  // Try to figure out what type of meeting we were in by scanning the text for keywords.
  // Note that not all types we might find are allowed in the curent opportunity stage.
  // When that happens, switch the stage so we can report on the indentified meeting type.
  // Search from most difficult to least difficult (bottom up).
  // More design info in build_se_events comments.
  // Returns the detected meeting type and required stage.
  
  let map = [];
  let rv = {};
  
  if (stage == "Discovery & Qualification") {
    
    
     
    // To detect default meeting type sections, set meeting: to "" (empty string is --None-- in SF)
    if (MEETINGS_HAVE_DEFAULT) {
      rv = { stage : "Discovery & Qualification", meeting : "Discovery"}; 
    }
    else {
      rv = { stage : "Discovery & Qualification", meeting : ""};
    }
    
    map = [
    // Although Salesforce says it will accept "Support" with Discovery & Qual stage,
    // not so much in practice as of May 2020. Moved the following from Support to PoV:
    // troubleshoot, support and issue.
    // Look for Discovery stuff first if we are in discovery stage
      {meeting : "Pilot", regex : /pilot/, stage : "Success Planning"},
      {meeting : "Health Check", regex : /health check/, stage : "Technical & Business Validation"},
      {meeting : "Standard Workshop", regex : /workshop/, stage : "Technical & Business Validation"},
      {meeting : "Product Deep Dive", regex : /deep dive/, stage : "Technical & Business Validation"},
      {meeting : "Controlled POV", regex : /implementation/, stage : "Technical & Business Validation"},
      {meeting : "Controlled POV", regex : /troubleshoot/, stage : "Technical & Business Validation"},
      {meeting : "Controlled POV", regex : /support/, stage : "Technical & Business Validation"},
      {meeting : "Controlled POV", regex : /issue/, stage : "Technical & Business Validation"},
      {meeting : "Controlled POV", regex : /pov/, stage : "Technical & Business Validation"},
      {meeting : "Controlled POV", regex : /poc/, stage : "Technical & Business Validation"},
      {meeting : "Shadow", regex : /shadow/, stage : "Discovery & Qualification"},
      {meeting : "Happy Hour", regex : /happy hour/, stage : "Discovery & Qualification"},
      {meeting : "Happy Hour", regex : /lunch/, stage : "Discovery & Qualification"},
      {meeting : "Happy Hour", regex : /coffee/, stage : "Discovery & Qualification"},
      {meeting : "Happy Hour", regex : /dinner/, stage : "Discovery & Qualification"},
      {meeting : "Happy Hour", regex : /drinks/, stage : "Discovery & Qualification"},
      {meeting : "Technical Office Hours", regex : /training/, stage : "Discovery & Qualification"},
      {meeting : "Technical Office Hours", regex : /setup/, stage : "Discovery & Qualification"},
      {meeting : "Technical Office Hours", regex : /setup/, stage : "Discovery & Qualification"},
      {meeting : "Product Overview", regex : /presentation/, stage : "Discovery & Qualification"},
      {meeting : "Product Overview", regex : /pitch/, stage : "Discovery & Qualification"},
      {meeting : "Product Overview", regex : /briefing/, stage : "Discovery & Qualification"},
      {meeting : "Product Overview", regex : /architecture/, stage : "Discovery & Qualification"},
      {meeting : "Product Overview", regex : /overview/, stage : "Discovery & Qualification"},
      {meeting : "Product Overview", regex : /whiteboard/, stage : "Discovery & Qualification"},
      {meeting : "Product Overview", regex : /roadmap/, stage : "Discovery & Qualification"},
      {meeting : "Discovery", regex : /planning/, stage : "Discovery & Qualification"},
      {meeting : "Discovery", regex : /discovery/, stage : "Discovery & Qualification"},
      {meeting : "Discovery", regex : /discussion/, stage : "Discovery & Qualification"},
      {meeting : "Discovery", regex : /touchpoint/, stage : "Discovery & Qualification"},
      {meeting : "Discovery", regex : /introduction/, stage : "Discovery & Qualification"},
      {meeting : "Discovery", regex : /sync/, stage : "Discovery & Qualification"},
      {meeting : "Discovery", regex : /review/, stage : "Discovery & Qualification"},
      {meeting : "Demo", regex : /demo/, stage : "Discovery & Qualification"}]; // want demo to take priority over POV
    
  }
  else if (stage == "Technical & Business Validation") {
    
      
    // To detect default meeting type sections, set meeting: to "" (empty string is --None-- in SF)
    if (MEETINGS_HAVE_DEFAULT) {
      rv = { stage : "Technical & Business Validation", meeting : "Technical Office Hours"}; 
    }
    else {
      rv = { stage : "Technical & Business Validation", meeting : ""};
    }
    
    map = [
      {meeting : "Pilot", regex : /pilot/, stage : "Success Planning"},
      {meeting : "Shadow", regex : /shadow/, stage : "Technical & Business Validation"},
      {meeting : "Happy Hour", regex : /happy hour/, stage : "Technical & Business Validation"},
      {meeting : "Happy Hour", regex : /lunch/, stage : "Technical & Business Validation"},
      {meeting : "Happy Hour", regex : /coffee/, stage : "Technical & Business Validation"},
      {meeting : "Happy Hour", regex : /dinner/, stage : "Technical & Business Validation"},
      {meeting : "Happy Hour", regex : /drinks/, stage : "Technical & Business Validation"},
      {meeting : "Product Overview", regex : /presentation/, stage : "Technical & Business Validation"},
      {meeting : "Product Overview", regex : /pitch/, stage : "Technical & Business Validation"},
      {meeting : "Product Overview", regex : /briefing/, stage : "Technical & Business Validation"},
      {meeting : "Product Overview", regex : /overview/, stage : "Technical & Business Validation"},
      {meeting : "Product Overview", regex : /whiteboard/, stage : "Technical & Business Validation"},
      {meeting : "Discovery", regex : /discovery/, stage : "Discovery & Qualification"},
      {meeting : "Discovery", regex : /introduction/, stage : "Discovery & Qualification"},
      {meeting : "Product Roadmap", regex : /roadmap/, stage : "Technical & Business Validation"},
      {meeting : "Technical Office Hours", regex : /training/, stage : "Technical & Business Validation"},
      {meeting : "Technical Office Hours", regex : /sync/, stage : "Technical & Business Validation"},
      {meeting : "Technical Office Hours", regex : /setup/, stage : "Technical & Business Validation"},
      {meeting : "Technical Office Hours", regex : /discussion/, stage : "Technical & Business Validation"},
      {meeting : "Technical Office Hours", regex : /touchpoint/, stage : "Technical & Business Validation"},
      {meeting : "Technical Office Hours", regex : /review/, stage : "Technical & Business Validation"},
      {meeting : "Technical Office Hours", regex : /support/, stage : "Technical & Business Validation"},
      {meeting : "Technical Office Hours", regex : /troubleshoot/, stage : "Technical & Business Validation"},
      {meeting : "Technical Office Hours", regex : /issue/, stage : "Technical & Business Validation"},
      {meeting : "Health Check", regex : /health check/, stage : "Technical & Business Validation"},
      {meeting : "Demo", regex : /demo/, stage : "Discovery & Qualification"},
      {meeting : "Standard Workshop", regex : /workshop/, stage : "Technical & Business Validation"},
      {meeting : "Product Deep Dive", regex : /deep dive/, stage : "Technical & Business Validation"},
      {meeting : "Product Deep Dive", regex : /architecture/, stage : "Technical & Business Validation"},
      {meeting : "Controlled POV", regex : /implementation/, stage : "Technical & Business Validation"},
      {meeting : "Controlled POV", regex : /pov/, stage : "Technical & Business Validation"},
      {meeting : "Controlled POV", regex : /poc/, stage : "Technical & Business Validation"}]; 
  }
  else if (stage == "Closed/Won") {
    
    // Salesforce will not accept Implementation or Support in Closed/Won stage. 
    // For now we will use "--None--" to indicate Post-Sales activity!
    if (MEETINGS_HAVE_DEFAULT) {
      rv = { stage : "Closed/Won", meeting : ""}; // For now this means post-sales
    }
    else {
      rv = { stage : "Closed/Won", meeting : ""};
    }
    
    // Salesforce will not accept Implementation or Support in Closed/Won stage. 
    // For now we will use "--None--" to indicate Post-Sales activity!
    map = [
      {meeting : "Shadow", regex : /shadow/, stage : "Closed/Won"},
      {meeting : "Customer Business Review", regex : /happy hour/, stage : "Closed/Won"},
      {meeting : "Customer Business Review", regex : /presentation/, stage : "Closed/Won"},
      {meeting : "Customer Business Review", regex : /pitch/, stage : "Closed/Won"},
      {meeting : "Customer Business Review", regex : /briefing/, stage : "Closed/Won"},
      {meeting : "Customer Business Review", regex : /overview/, stage : "Closed/Won"},
      {meeting : "Customer Business Review", regex : /kick/, stage : "Closed/Won"},  // kick-offs Cadence  
      {meeting : "Customer Business Review", regex : /cadence/, stage : "Closed/Won"}, 
      {meeting : "Customer Business Review", regex : /touchpoint/, stage : "Closed/Won"},
      {meeting : "Customer Business Review", regex : /coffee/, stage : "Closed/Won"},
      {meeting : "Customer Business Review", regex : /discussion/, stage : "Closed/Won"},
      {meeting : "Discovery", regex : /discovery/, stage : "None"}, // Have to push stage to None for Discovery
      {meeting : "Product Roadmap", regex : /roadmap/, stage : "Closed/Won"},
      {meeting : "Training", regex : /training/, stage : "Closed/Won"},
      {meeting : "Training", regex : /sync/, stage : "Closed/Won"},
      {meeting : "Training", regex : /setup/, stage : "Closed/Won"},
      {meeting : "Training", regex : /workshop/, stage : "Closed/Won"},
      {meeting : "Training", regex : /deep dive/, stage : "Closed/Won"},
      {meeting : "Training", regex : /pov/, stage : "Closed/Won"},
      {meeting : "Customer Business Review", regex : /health check/, stage : "Closed/Won"},
      {meeting : "Customer Business Review", regex : /demo/, stage : "Closed/Won"},
      {meeting : "", regex : /support/, stage : "Closed/Won"},
      {meeting : "", regex : /implementation/, stage : "Closed/Won"},
      {meeting : "", regex : /troubleshoot/, stage : "Closed/Won"},
      {meeting : "", regex : /help/, stage : "Closed/Won"},
      {meeting : "", regex : /issue/, stage : "Closed/Won"},
      {meeting : "", regex : /pilot/, stage : "Closed/Won"}];
      // This is suppossed to work, but doesn't as of May 2020
      /*
      {meeting : "Implementation", regex : /support/, stage : "Closed/Won"},
      {meeting : "Implementation", regex : /implementation/, stage : "Closed/Won"},
      {meeting : "Implementation", regex : /troubleshoot/, stage : "Closed/Won"},
      {meeting : "Implementation", regex : /help/, stage : "Closed/Won"},
      {meeting : "Implementation", regex : /issue/, stage : "Closed/Won"},
      {meeting : "Implementation", regex : /pilot/, stage : "Closed/Won"}];
      */
  }
  else {
    rv = { stage : "Closed/Lost", meeting : ""};
    return rv;
  }
  
  let x = text.toLowerCase();
  
  for (i = map.length - 1; i >= 0; i--) { 
    if (map[i].regex.test(x)) {
      rv.stage = map[i].stage;
      rv.meeting = map[i].meeting;
      break;
    }
  }
  
  return rv;
  
}

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

function filterAndAnalyzeDescription_(text) {

  let rv = {hasTeleconference : false, filteredText : text, prepTime : 0}

  if (!text) return rv;
  
  // 
  // For the most part, assume everything after a Zoom, Webex, getclockwise, Microsoft Team Meeting
  // intros, ... to the end is garbage. (No one will add important info below all that, intentionally at least)
  //

  // Process Prep tag
  let regex = /[Pp][Rr][Ee][Pp] *: *[0-9]+[ ]?[MmHhDd]?/; // e.g. Prep:30m
  let prepArray = text.match(regex);
  if (prepArray && prepArray[0]) {
    let kv = prepArray[0].split(':');
    rv.prepTime = parseInt(kv[1]); // minutes
    let units = kv[1].match(/[MmHhDd]$/);
    if (units) {
      if (units == "h" || units == "H") {
        rv.prepTime = rv.prepTime * 60;
      }
      else if (units == "d" || units == "D") {
        rv.prepTime = rv.prepTime * 60 * 8; // 8 hour workday
      }
    }
  }
    
  
  if (text.indexOf("\─\─\─\─\─\─\─\─\─\─[\s\S]*Join Zoom Meeting") != -1) {
    regex = /\─\─\─\─\─\─\─\─\─\─[\s\S]*Join Zoom Meeting[\s\S]*\─\─\─\─\─\─\─\─\─\─/;
    rv.hasTeleconference = true;
  }
  else if (text.indexOf("Join Zoom Meeting") != -1) {
    regex = /Join Zoom Meeting[\s\S]*/;
    rv.hasTeleconference = true;
  }
  else if (text.indexOf("Do not delete or change any of the following text.") != -1) {
    regex = /Do not delete or change any of the following text.[\s\S]*/;
    rv.hasTeleconference = true;
  }
  else if (text.indexOf("Join Microsoft Teams Meeting") != -1) {
    regex = /Join Microsoft Teams Meeting[\s\S]*/;
    rv.hasTeleconference = true;
  }
  else if (text.indexOf("getclockwise.com") != -1) {
    regex = /getclockwise.com[\s\S]*/;
    rv.hasTeleconference = true;
  }
  
  if (rv.hasTeleconference) {
    text = text.replace(regex, "");
  }
  
  if (text.indexOf("zoom.us") != -1) {
    rv.hasTeleconference = true;
  }
  rv.filteredText = text;
  return rv;
}

function isRepPresent_(createdBy, attendees) {
  if (staffEmailToRoleMap[createdBy] && staffEmailToRoleMap[createdBy] == STAFF_ROLE_REP) {
    return true;
  }
  
  for (let j = 0; j < attendees.length; j++) {
    let attendeeEmail = attendees[j].trim();    
    if (staffEmailToRoleMap[attendeeEmail] && staffEmailToRoleMap[attendeeEmail] == STAFF_ROLE_REP) {
      return true;
    }
  }
  
  return false;
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

function clearTab_(tab_name, header) {
  
  if (!tab_name || !header) return;
  
  let sheet = SpreadsheetApp.getActive().getSheetByName(tab_name);
  sheet.clearContents();  
  let range = sheet.getRange(1,1,1,header[0].length);
  range.setValues(header);
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
    
    rowOffset++;
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
  sheet = SpreadsheetApp.getActive().getSheetByName("Expanded");
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
      if (accountLog[domain]) {
        Logger.log("WARNING: Found account with a duplicate email domain - " + accountInfo[j][accountName] + ":" + domain);
        continue; // Take the first
      }
      accountLog[domain] = true;
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
  // Load Opportunity's Stage History
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
  // Convert calendar invites (Calendart tab) to SE events (Events tab)
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
    // Manual Overrides!
    //
    
    //DAK
     
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
