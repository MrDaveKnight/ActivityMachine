// Activity Machine
//
// This is a tool to automatically create SE Activities (Field Events in Salesforce) from calendar invites
// that include customers, partners or leads. Internal-only meetings are not included. All customer/partner/lead facing
// non-recurring invites are included. All customer/partner/lead facing recurring invites that have been "accepted" 
// (SE has indicated they will be going) are also included.
// 
// The tool identifies the account(s) associated with each invite, along with the opportunity if any, based on the 
// attendees and the products being discussed. The meeting type is derived from keywords in the invite
// such as "demo" or "pov". Reference the lookForMeetingType_() method for more detail.
//
// This script provides logic to -
// 1. Import google calendar invites into the google sheet
// 2. Generate SE Activities from those invites, saving them in a "staging" tab in the google sheet
// 3. Upload SE Activities; moving staged SE Activities into the "Upload" tab for Zap upload
// 4. Generate "expanded" SE Activities with machine IDs replaced by names for easier review
// 5. Import "missing" customers and leads (customers that are out-of-region, and leads for attendees without an account)
// 6. Clear various tabs and logs
//
// A Zapier zap is configured to watch for SE Activities that appear in the Upload tab, and push them up 
// into Salesforce. Two more zaps are configured to watch for "missing" domains and emails in order to import
// out-of-region customers and potential leads
// 
// ** Remember to turn off the associated Zap(s) before you delete records from the Upload or "Missing" tabs. **
//
// Information from Salesforce needed by the invite-to-event transform logic is stored in other google sheet tabs,
// and imported via "Dataconnector for Salesforce".
//
// Written by Dave Knight, Rez Dogs 2020, knight@hashicorp.com

// Clear meeting type default
const MEETINGS_HAVE_DEFAULT=true;

// Information for accounts, staff and opportunities is imported directly from
// Salesforce and saved in spreadsheet tabs.
// Calendar invite information has been imported directly from google calendar into a tab.
// The indexes below identify the column locations of the imported data.

const RUN_PARMS = "Run Settings"

const MISSING_DOMAINS = "Missing Domains (Zap)";
const MISSING_EMAILS = "Missing Emails (Zap)";
const MISSING_EMAIL_DOMAIN = 0;

const IN_REGION_CUSTOMERS = "In Region Customers"; // Our Region's Customers or Prospects - Accounts for short. They are accounts in Salesforce
const MISSING_CUSTOMERS = "Missing Customers (Zap)"; // Dynamic account imports for each missing customer
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

const MISSING_LEADS = "Missing Leads (Zap)";
const LEAD_COLUMNS = 3;
const LEAD_ID = 0;
const LEAD_NAME =1;
const LEAD_EMAIL = 2;
const LEAD_HEADER = [['18-Digit Lead ID', 'Account Name', 'Email']];			

const STAFF = "Staff"
const STAFF_COLUMNS = 5;
const STAFF_ID =3;
const STAFF_NAME = 0;
const STAFF_EMAIL = 1;
const STAFF_ROLE = 4;
const STAFF_ROLE_REP = "EAM 1"

const OPPORTUNITIES = "Opportunities";
const OP_COLUMNS = 12;
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
const OP_ACTIVITY_DATE = 12;

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
const EVENTS_EXPANDED = "Expanded"
const EVENT_COLUMNS = 14;
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
const EVENT_QUALITY = 12;
const EVENT_LEAD = 13;
const EVENT_HEADER = [["Assigned To", "Opportunity Stage", "Meeting Type", "Related To", "Subject", "Start", "End", "Rep Attended", "Primary Product", "Description", "Logistics", "Prep", "Quality", "Lead"]]

// Global log
const LOG_TAB = "Log";
var AM_LOG; 
var AM_LOG_ROW; 

var emailToCustomerMap = {};
var emailToPartnerMap = {};
var emailToLeadMap = {};
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
const LEAD_TYPE = 4;
var accountType = {}; // Index by account ID

var missingAccounts = {}; // For tracking missing account data in lookForAccounts_()
var potentialLeadEmails = {}; // For finding Leads should a missing account not exist.

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
                  .addItem('Log Tab', 'menuItem12_')
                  .addSeparator()
                  .addItem('Calendar Tab', 'menuItem8_')
                  .addSeparator()
                  .addItem('Event Tab', 'menuItem9_')
                  .addSeparator()
                  .addItem('Expanded Tab', 'menuItem11_')
                  .addSeparator()
                  .addItem('Missing Accounts Tabs (Turn off Zapier!)', 'menuItem5_')
                  .addSeparator()
                  .addItem('Upload Tab (Turn off Zapier!)', 'menuItem6_'))
      .addToUi();
}