// Activity Machine
//
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
// 4. Generate "unveiled" SE Activities with machine IDs replaced by names for easier review
// 5. Clear various tabs and logs
//
// A Zapier zap is configured to watch for SE Activities that appear in the Upload tab, and push them up 
// into Salesforce.
// 
// ** Remember to turn off the associated Zap(s) before you delete records from the Upload tab. **
//
// Information from Salesforce needed by the invite-to-event transform logic is stored in other google sheet tabs,
// and imported via "Dataconnector for Salesforce".
//
// Written by Dave Knight, Rez Dogs 2020, knight@hashicorp.com

const GAS_VERSION_STRING = "1.2.1";
const GAS_VERSION = 121; 
const MIN_SCHEMA_VERSION = 20; // 2.0

// Clear meeting type default
const MEETINGS_HAVE_DEFAULT=true;

// Information for accounts, staff and opportunities is imported directly from
// Salesforce and saved in spreadsheet tabs.
// Calendar invite information has been imported directly from google calendar into a tab.
// The indexes below identify the column locations of the imported data.

const RUN_PARMS = "Run Settings"
const CHOICES = "Choices"

// Event types
const OP_EVENT = "Opportunity";
const CUSTOMER_EVENT = "Customer";
const PARTNER_EVENT = "Partner";
const LEAD_EVENT = "Lead";

// Filter tables to help us deal with massive amounts of Salesforce records
const MISSING_DOMAINS = "Missing Domains";
const MISSING_CUSTOMERS = "Missing Customers";
const POTENTIAL_LEADS = "Lead Accounts";
const IN_PLAY_CUSTOMERS = "In Play Customers";
const IN_PLAY_PARTNERS = "In Play Partners";
const FILTER_EMAIL_DOMAIN = 0; // This is actually the full email of the lead (need to change name)
const FILTER_ACCOUNT_ID = 0;
const FILTER_TYPE_DOMAIN = 1;
const FILTER_TYPE_ID = 2;

const IN_REGION_CUSTOMERS = "In-Region Customers"; // Our Region's Customers or Prospects; Accounts for short. They are accounts in Salesforce
const ALL_CUSTOMERS = "Customers"; // ALL Customers and Prospects in Salesforce! A big list. Logic handles I/O to this tab in chunks and target filters.
const CUSTOMER_COLUMNS = 7;
const CUSTOMER_ID = 1;
const CUSTOMER_NAME = 3;
const CUSTOMER_EMAIL_DOMAIN = 5;
const CUSTOMER_ALT_EMAIL_DOMAINS = 6;
const CUSTOMER_HEADER = [['Account Owner', '18-Digit Account ID', 'Solutions Engineer', 'Account Name', 'Owner Region', 'Email Domain', 'Alt Email Domains']];		

const PARTNERS = "Partners"; // ALL Partners in Salesforce, also an account btw. Relatively long list. Logic handles I/O to this tab in chunks.
const PARTNER_COLUMNS = 4;
const PARTNER_ID = 1;
const PARTNER_NAME =2;
const PARTNER_EMAIL_DOMAIN = 3;
const PARTNER_ALT_EMAIL_DOMAINS = 4;
const PARTNER_HEADER = [['Account Owner', '18-Digit Account ID', 'Account Name', 'Email Domain']];			// FIXME Alt email			

const LEADS = "Leads";  // // ALL Leads in Salesforce! A huge list! 150K records at last count. Logic handles I/O to this tab in chunks and target filters.
const LEAD_COLUMNS = 3;
const LEAD_ID = 0;
const LEAD_NAME =1;
const LEAD_EMAIL = 2;
const LEAD_HEADER = [['18-Digit Lead ID', 'Account Name', 'Email']];			

const STAFF = "Staff"
const STAFF_COLUMNS = 6;
const STAFF_ID = 3;
const STAFF_NAME = 0;
const STAFF_EMAIL = 1;
const STAFF_ROLE = 4;
const STAFF_LONG_ID = 5;
const STAFF_ROLE_REPS = "E1:G1:E2:C1:CAM"; // All the "Job Roles" for reps
const STAAFF_HEADER = [['18-Digit Lead ID', 'Account Name', 'Email']];	

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
const EVENT_COLUMNS = 23; // If you add one, update this too
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
const EVENT_NOTES = 14;
// Following don't go up to Salesforce
const EVENT_ACCOUNT_TYPE = 15;
const EVENT_PROCESS = 16;
const EVENT_RECURRING = 17;
const EVENT_DEMO = 18;
const EVENT_POV = 19;
const EVENT_WORKSHOP = 20;
const EVENT_DIVE = 21;
const EVENT_ATTENDEES = 22;
const EVENT_HEADER = [["Assigned To", "Opportunity Stage", "Meeting Type", "Related To", "Subject", "Start", "End", "Rep Attended", "Primary Product", "Description", "Logistics", "Prep", "Quality", "Lead", "Notes", "Account Type", "Process", "Recurring", "Demo", "POV", "Workshop", "Deep Dive", "Attendees"]];
const UPLOAD_STAGE_HEADER = [["Assigned To", "Opportunity Stage", "Meeting Type", "Related To", "Subject", "Start", "End", "Rep Attended", "Primary Product", "Description", "Logistics", "Prep", "Quality", "Lead", "Notes"]];

const EVENTS_UNVEILED = "Review"
const REVIEW_COLUMNS = 18;
const REVIEW_ASSIGNED_TO = 0;
const REVIEW_EVENT_TYPE = 1; // Must be before REVIEW_RELATED_TO and REVIEW_LEAD or you will break the menu
const REVIEW_RELATED_TO = 2; // Opportunity ID or Account ID
const REVIEW_OP_STAGE = 3;
const REVIEW_START = 4;
const REVIEW_END = 5;
const REVIEW_SUBJECT = 6;
const REVIEW_PRODUCT = 7;
const REVIEW_DESC = 8;
const REVIEW_MEETING_TYPE = 9;
const REVIEW_REP_ATTENDED = 10;
const REVIEW_LOGISTICS = 11;
const REVIEW_PREP_TIME = 12;
const REVIEW_QUALITY = 13;
const REVIEW_LEAD = 14;
const REVIEW_NOTES = 15;
const REVIEW_ACCOUNT_TYPE = 16;
const REVIEW_PROCESS = 17;
const REVIEW_HEADER = [["Assigned To", "Event Type", "Related To", "Opportunity Stage", "Start", "End", "Subject", "Primary Product", "Description", "Meeting Type", "Rep Attended", "Logistics", "Prep", "Quality", "Lead", "Notes", "Account Type", "Process"]]
// There are field protections setup in unveil_se_events that are hardcoded to this header. If you change this, make sure the protections in unveil_se_events are correct or updated.

// Record original values for highlighting changes by the reviewers
// No header. Not meant to be seen by a user
const REVIEW_ORIG_EVENT_TYPE = 31;
const REVIEW_ORIG_RELATED_TO = 32; // Opportunity ID or Account ID
const REVIEW_ORIG_OP_STAGE = 33;
const REVIEW_ORIG_PRODUCT = 34;
const REVIEW_ORIG_DESC = 35;
const REVIEW_ORIG_MEETING_TYPE = 36;
const REVIEW_ORIG_REP_ATTENDED = 37;
const REVIEW_ORIG_LOGISTICS = 38;
const REVIEW_ORIG_PREP_TIME = 39;
const REVIEW_ORIG_QUALITY = 40;
const REVIEW_ORIG_LEAD = 41;
const REVIEW_ORIG_NOTES = 42;
const REVIEW_ORIG_ACCOUNT_TYPE = 43;
const REVIEW_ORIG_PROCESS = 44;
var inReviewInitialization = false; // flag to disable onEdit functionality while the tab is being built

// Choice Tables for Review Tab
// The Saleforce configuration tables are too massive for field dropdown selection.
// Use these tables to filter things down.
// Give all the domains in the calendar invites, include ALL accounts (even duplicates)
// Give all the accounts, include all possible opportunities
// A filtered selection of leads already exists in Lead Accounts
const CHOICE_ACCOUNT = "Account Choices";
const CHOICE_ACCOUNT_NAME = 0;
const CHOICE_ACCOUNT_ID = 1;
const CHOICE_PARTNER = "Partner Choices";
const CHOICE_PARTNER_NAME = 0;
const CHOICE_PARTNER_ID = 1;
const CHOICE_OP = "Op Choices";
const CHOICE_OP_NAME = 0;
const CHOICE_OP_ID = 1;
const CHOICE_LEAD = "Lead Choices"
const CHOICE_LEAD_EMAIL = 0;

const MAX_ROWS = 2000; // Use for warnings and data validation limits


// Global log
const LOG_TAB = "Log";
var AM_LOG; 
var AM_LOG_ROW; 

const INTERNAL_CUSTOMER_TYPE = 1;
const EXTERNAL_CUSTOMER_TYPE = 2;
const PARTNER_TYPE = 3;
const LEAD_TYPE = 4;
var accountTypeG = {}; // Index by account ID
var accountTypeNamesG = [
  "N/A",
  "In-region Customer",
  "Out-or-region Customer",
  "Partner",
  "Lead"];
var emailToCustomerMapG = {}; // Email domain to customer account id
var emailToPartnerMapG = {}; // Email domain to customer account id
var emailToLeadMapG = {};     // Full email address to lead id (different from customers and partners!)
var staffEmailToIdMapG = {};
var staffEmailToRoleMapG = {};
var staffNameToEmailMapG = {}; // Looks like I don't use this anywhere yet
var opByCustomerAndProductG = {};
var numberOfOpsByCustomerG = {};
var primaryOpByCustomerG = {};
var primaryProductByOpG = {};
var stageMilestonesByOpG = {};
var accountTeamByOpG = {};

// Review tab choice list trackers - account name:id
var duplicateAccountsG = {}; // All non-primary accounts that share a domain with the primary
var duplicatePartnersG = {}; 
var primaryAccountsG = {};
var primaryPartnersG = {};

var testmeG = {};


// Product codes for table index keys
const TERRAFORM = "T";
const VAULT = "V";
const CONSUL = "C";
const NOMAD = "N";

// 
// Stats
//

var statsLedgerG = { global: { agendaItems: { demo: 0, pov: 0, workshop: 0, deepdive: 0 } }, users : {} }; // see the collectStats_ function in eutility.gs
var statsOutputWorksheetIdG = null;
const MEETING = "Meetings";
const MEETING_COUNTS_ORIGIN_ROW = 3;
const MEETING_COUNTS_ORIGIN_COL = 2;


// Track special prep invites
// Subject contains Prep:<Subject of meeting the prep is for>
const PREP = "prep";
var prepCalendarEntries = {};

const PROCESS_UPLOAD = "Upload";
const PROCESS_SKIP = "Skip";

// Debug tracing
const IS_TRACE_ACCOUNT_ON = false;
const TRACE_ACCOUNT_ID = "0011C00001rtkI3"
const TRACE_ACCOUNT_ID_LONG = "0011C00001rtkI3QAI"

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Activity')
      .addItem('Process Events', 'menuItem3_')
      .addSeparator()
      .addSubMenu(ui.createMenu('Process Events by Stage')
          .addItem('Import Calendars', 'menuItem1_')
          .addSeparator()
          .addItem('Generate Events', 'menuItem2_')
          .addSeparator()
          .addItem('Unveil Events', 'menuItem10_')
          .addSeparator()
          .addItem('Generate & Unveil Events', 'menuItem7_'))
      .addSeparator()
      .addItem('Reconcile Events', 'menuItem13_')
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
                  .addItem('Review Tab', 'menuItem11_')
                  .addSeparator()
                  .addItem('Upload Tab (Turn off Zapier!)', 'menuItem6_')
                  .addSeparator()
                  .addItem('Everything (Turn off Zapier!)', 'menuItem14_')
                  .addSeparator()
                  .addItem('Configuration', 'menuItem15_'))
      .addSeparator()
      .addSubMenu(ui.createMenu('Help')
                  .addItem('Process Events', 'menuItem24_')
                  .addSeparator()
                  .addItem('Import Calendars', 'menuItem20_')
                  .addSeparator()
                  .addItem('Generate Events', 'menuItem22_')
                  .addSeparator()
                  .addItem('Unveil Events', 'menuItem23_')
                  .addSeparator()
                  .addItem('Reconcile Events', 'menuItem26_')
                  .addSeparator()
                  .addItem('Upload Events', 'menuItem25_')
                  .addSeparator()
                  .addItem('Check Version', 'menuItem27_'))
      .addSeparator()
      .addItem('Check Version', 'menuItem30_')
      .addToUi();
}