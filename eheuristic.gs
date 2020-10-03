
const STAGE_D = "Discovery & Qualification";
const STAGE_V = "Technical & Business Validation";
const STAGE_SP = "Success Planning";
const STAGE_C = "Closed/Won";
const STAGE_N = "None";

// Tests to make sure the Salesforce API doesn't reject our Meeting-Type:Op-Stage selections
var validStagesByMeeting = {};
validStagesByMeeting["Demo" + STAGE_D] = true;
validStagesByMeeting["Product Deep Dive" + STAGE_V] = true;
validStagesByMeeting["Standard Workshop" + STAGE_V] = true;
validStagesByMeeting["Health Check" + STAGE_V] = true;
validStagesByMeeting["Product Overview" + STAGE_D] = true;
validStagesByMeeting["Product Overview" + STAGE_V] = true;
validStagesByMeeting["Controlled POV" + STAGE_V] = true;
validStagesByMeeting["Discovery" + STAGE_D] = true;
validStagesByMeeting["Technical Office Hours" + STAGE_D] = true;
validStagesByMeeting["Technical Office Hours" + STAGE_V] = true;
validStagesByMeeting["Happy Hour" + STAGE_D] = true;
validStagesByMeeting["Happy Hour" + STAGE_V] = true;
validStagesByMeeting["Shadow" + STAGE_D] = true;
validStagesByMeeting["Shadow" + STAGE_V] = true;
validStagesByMeeting["Shadow" + STAGE_C] = true;
validStagesByMeeting["Pilot" + STAGE_SP] = true;
validStagesByMeeting["Product Roadmap" + STAGE_V] = true;
validStagesByMeeting["Product Roadmap" + STAGE_C] = true;
validStagesByMeeting["Customer Business Review" + STAGE_C] = true;
validStagesByMeeting["Training" + STAGE_C] = true;

validStagesByMeeting["Demo" + STAGE_N] = true;
validStagesByMeeting["Product Deep Dive" + STAGE_N] = true;
validStagesByMeeting["Standard Workshop" + STAGE_N] = true;
validStagesByMeeting["Health Check" + STAGE_N] = true;
validStagesByMeeting["Product Overview" + STAGE_N] = true;
validStagesByMeeting["Controlled POV" + STAGE_N] = true;
validStagesByMeeting["Discovery" + STAGE_N] = true;
validStagesByMeeting["Technical Office Hours" + STAGE_N] = true;
validStagesByMeeting["Happy Hour" + STAGE_N] = true;
validStagesByMeeting["Shadow" + STAGE_N] = true;
validStagesByMeeting["Pilot" + STAGE_N] = true;
validStagesByMeeting["Product Roadmap" + STAGE_N] = true;
validStagesByMeeting["Training" + STAGE_N] = true;

// Overrides to make sure the Salesforce API doesn't reject our Meeting-Type:Op-Stage selections
var defaultStageForMeeting = {
  Demo : STAGE_D,
  "Product Deep Dive" : STAGE_V,
  "Standard Workshop" : STAGE_V,
  "Health Check" : STAGE_V,
  "Product Overview" : STAGE_D,
  "Controlled POV" : STAGE_V,
  Discovery : STAGE_D,
  "Technical Office Hours" : STAGE_D,
  "Happy Hour" : STAGE_V,
  Shadow : STAGE_D,
  Pilot : STAGE_SP,
  "Product Roadmap" : STAGE_V,
  "Customer Business Review" : STAGE_C,
  Training : STAGE_C};

/* Don't think I need to use this for UI data validation choices, at least not yet.
var stageValidation = {
  Demo : STAGE_D,
  "Product Deep Dive" : STAGE_V,
  "Standard Workshop" : STAGE_V,
  "Health Check" : STAGE_V,
  "Product Overview" : STAGE_D + "," + STAGE_V,
  "Controlled POV" : STAGE_V,
  Discovery : STAGE_D + "," + "None",
  "Technical Office Hours" : STAGE_D + "," + STAGE_V,
  "Happy Hour" : STAGE_D + "," + STAGE_V,
  Shadow : STAGE_D + "," + STAGE_V + "," + STAGE_C,
  Pilot : STAGE_SP,
  "Product Roadmap" : STAGE_V + "," + STAGE_C,
  "Customer Business Review" : STAGE_C,
  Training : STAGE_C};
  */



 

var discoveryMap = [
  // Although Salesforce says it will accept "Support" with Discovery & Qual stage,
  // not so much in practice as of May 2020. Moved the following from Support to PoV:
  // troubleshoot, support and issue.
  // Look for Discovery stuff first if we are in discovery stage
  {meeting : "Pilot", regex : /pilot/, stage : STAGE_SP},
  {meeting : "Controlled POV", regex : /implementation/, stage : STAGE_V},
  {meeting : "Controlled POV", regex : /troubleshoot/, stage : STAGE_V},
  {meeting : "Controlled POV", regex : /support/, stage : STAGE_V},
  {meeting : "Controlled POV", regex : /issue/, stage : STAGE_V},
  {meeting : "Shadow", regex : /shadow/, stage : STAGE_D},
  {meeting : "Happy Hour", regex : /happy hour/, stage : STAGE_D},
  {meeting : "Happy Hour", regex : /lunch/, stage : STAGE_D},
  {meeting : "Happy Hour", regex : /coffee/, stage : STAGE_D},
  {meeting : "Happy Hour", regex : /dinner/, stage : STAGE_D},
  {meeting : "Happy Hour", regex : /drinks/, stage : STAGE_D},
  {meeting : "Technical Office Hours", regex : /training/, stage : STAGE_D},
  {meeting : "Technical Office Hours", regex : /setup/, stage : STAGE_D},
  {meeting : "Technical Office Hours", regex : /setup/, stage : STAGE_D},
  {meeting : "Discovery", regex : /planning/, stage : STAGE_D},
  {meeting : "Discovery", regex : /discovery/, stage : STAGE_D},
  {meeting : "Discovery", regex : /discussion/, stage : STAGE_D},
  {meeting : "Discovery", regex : /touchpoint/, stage : STAGE_D},
  {meeting : "Discovery", regex : /introduction/, stage : STAGE_D},
  {meeting : "Discovery", regex : /sync/, stage : STAGE_D},
  {meeting : "Discovery", regex : /review/, stage : STAGE_D},
  {meeting : "Product Overview", regex : /presentation/, stage : STAGE_D},
  {meeting : "Product Overview", regex : /pitch/, stage : STAGE_D},
  {meeting : "Product Overview", regex : /briefing/, stage : STAGE_D},
  {meeting : "Product Overview", regex : /architecture/, stage : STAGE_D},
  {meeting : "Product Overview", regex : /overview/, stage : STAGE_D},
  {meeting : "Product Overview", regex : /whiteboard/, stage : STAGE_D},
  {meeting : "Controlled POV", regex : /pov/, stage : STAGE_V},
  {meeting : "Controlled POV", regex : /poc/, stage : STAGE_V},
  {meeting : "Product Overview", regex : /roadmap/, stage : STAGE_D},
  {meeting : "Health Check", regex : /health check/, stage : STAGE_V},
  {meeting : "Standard Workshop", regex : /workshop/, stage : STAGE_V},
  {meeting : "Product Deep Dive", regex : /deep dive/, stage : STAGE_V},
  {meeting : "Demo", regex : /demo/, stage : STAGE_D}]; // want demo to take priority over POV

var validationMap = [
  {meeting : "Pilot", regex : /pilot/, stage : STAGE_SP},
  {meeting : "Shadow", regex : /shadow/, stage : STAGE_V},
  {meeting : "Happy Hour", regex : /happy hour/, stage : STAGE_V},
  {meeting : "Happy Hour", regex : /lunch/, stage : STAGE_V},
  {meeting : "Happy Hour", regex : /coffee/, stage : STAGE_V},
  {meeting : "Happy Hour", regex : /dinner/, stage : STAGE_V},
  {meeting : "Happy Hour", regex : /drinks/, stage : STAGE_V},
  {meeting : "Product Overview", regex : /presentation/, stage : STAGE_V},
  {meeting : "Product Overview", regex : /pitch/, stage : STAGE_V},
  {meeting : "Product Overview", regex : /briefing/, stage : STAGE_V},
  {meeting : "Product Overview", regex : /overview/, stage : STAGE_V},
  {meeting : "Product Overview", regex : /whiteboard/, stage : STAGE_V},
  {meeting : "Discovery", regex : /discovery/, stage : STAGE_D},
  {meeting : "Discovery", regex : /introduction/, stage : STAGE_D},
  {meeting : "Product Roadmap", regex : /roadmap/, stage : STAGE_V},
  {meeting : "Technical Office Hours", regex : /training/, stage : STAGE_V},
  {meeting : "Technical Office Hours", regex : /sync/, stage : STAGE_V},
  {meeting : "Technical Office Hours", regex : /setup/, stage : STAGE_V},
  {meeting : "Technical Office Hours", regex : /discussion/, stage : STAGE_V},
  {meeting : "Technical Office Hours", regex : /touchpoint/, stage : STAGE_V},
  {meeting : "Technical Office Hours", regex : /review/, stage : STAGE_V},
  {meeting : "Technical Office Hours", regex : /support/, stage : STAGE_V},
  {meeting : "Technical Office Hours", regex : /troubleshoot/, stage : STAGE_V},
  {meeting : "Technical Office Hours", regex : /issue/, stage : STAGE_V},
  {meeting : "Health Check", regex : /health check/, stage : STAGE_V},
  {meeting : "Standard Workshop", regex : /workshop/, stage : STAGE_V},
  {meeting : "Product Deep Dive", regex : /deep dive/, stage : STAGE_V},
  {meeting : "Product Deep Dive", regex : /architecture/, stage : STAGE_V},
  {meeting : "Controlled POV", regex : /implementation/, stage : STAGE_V},
  {meeting : "Controlled POV", regex : /pov/, stage : STAGE_V},
  {meeting : "Controlled POV", regex : /poc/, stage : STAGE_V},
  {meeting : "Demo", regex : /demo/, stage : STAGE_D}]; 

var  closedMap = [
  {meeting : "Shadow", regex : /shadow/, stage : STAGE_C},
  {meeting : "Customer Business Review", regex : /happy hour/, stage : STAGE_C},
  {meeting : "Customer Business Review", regex : /presentation/, stage : STAGE_C},
  {meeting : "Customer Business Review", regex : /pitch/, stage : STAGE_C},
  {meeting : "Customer Business Review", regex : /briefing/, stage : STAGE_C},
  {meeting : "Customer Business Review", regex : /overview/, stage : STAGE_C},
  {meeting : "Customer Business Review", regex : /kick/, stage : STAGE_C},  // kick-offs Cadence  
  {meeting : "Customer Business Review", regex : /cadence/, stage : STAGE_C}, 
  {meeting : "Customer Business Review", regex : /touchpoint/, stage : STAGE_C},
  {meeting : "Customer Business Review", regex : /coffee/, stage : STAGE_C},
  {meeting : "Customer Business Review", regex : /discussion/, stage : STAGE_C},
  {meeting : "Discovery", regex : /discovery/, stage : "None"}, // Have to push stage to None for Discovery
  {meeting : "Product Roadmap", regex : /roadmap/, stage : STAGE_C},
  {meeting : "Training", regex : /training/, stage : STAGE_C},
  {meeting : "Training", regex : /sync/, stage : STAGE_C},
  {meeting : "Training", regex : /setup/, stage : STAGE_C},
  {meeting : "Training", regex : /workshop/, stage : STAGE_C},
  {meeting : "Training", regex : /deep dive/, stage : STAGE_C},
  {meeting : "Training", regex : /pov/, stage : STAGE_C},
  {meeting : "Customer Business Review", regex : /health check/, stage : STAGE_C},
  {meeting : "Customer Business Review", regex : /demo/, stage : STAGE_C},
  {meeting : "", regex : /support/, stage : STAGE_C},
  {meeting : "", regex : /implementation/, stage : STAGE_C},
  {meeting : "", regex : /troubleshoot/, stage : STAGE_C},
  {meeting : "", regex : /help/, stage : STAGE_C},
  {meeting : "", regex : /issue/, stage : STAGE_C},
  {meeting : "", regex : /pilot/, stage : STAGE_C}];
// This is suppossed to work, but doesn't as of May 2020
/*
{meeting : "Implementation", regex : /support/, stage : STAGE_C},
{meeting : "Implementation", regex : /implementation/, stage : STAGE_C},
{meeting : "Implementation", regex : /troubleshoot/, stage : STAGE_C},
{meeting : "Implementation", regex : /help/, stage : STAGE_C},
{meeting : "Implementation", regex : /issue/, stage : STAGE_C},
{meeting : "Implementation", regex : /pilot/, stage : STAGE_C}];
*/


/*      
var validStagesForMeetings = {
 
}
*/

function lookForAccounts_(attendees, customerMap, partnerMap) {

  // Scan email domains of attendees looking for accounts.

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
  rv.stats.others = 0; // Number of unidentified attendees

  
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
    hasNomad : false,
    count : function () {
      let retval = 0;
      if (this.hasTerraform) retval++;
      if (this.hasVault) retval++;
      if (this.hasNomad) retval++;
      if (this.hasConsul) retval++;
      return retval;
    },
    has : function (product) {
      let retVal = false;
      switch (product) {
        case "Terraform":
          retVal = this.hasTerraform;
          break;
        case "Vault":
          retVal = this.hasVault;
          break;
        case "Consul":
          retVal = this.hasConsul;
          break;
        case "Nomad":
          retVal = this.hasNomad;
          break;
        default:
          retVal = false;
      }
    },
    getOne : function () {
      let retval = "N/A";
      if (this.hasTerraform) { retval = "Terraform"; }
      else if (this.hasVault) { retval = "Vault"; }
      else if (this.hasConsul) { retval = "Consul"; }
      else if (this.hasNomad) { retval = "Nomad"; }
      return retval;
    }
  }
  
  if (!text) return returnValue;
  var x = text.toLowerCase();
  
  returnValue.hasTerraform = x.indexOf("terraform") != -1 || x.indexOf("tf cloud") != -1;
  if (!returnValue.hasTerraform) {
    let regex = RegExp("p?tf[ce]");
    returnValue.hasTerraform = regex.test(x);
  }
  
  returnValue.hasVault = x.indexOf("vault") != -1 || x.indexOf("secret") != -1 || x.indexOf("pki") != -1;
  
  returnValue.hasConsul = x.indexOf("consul") != -1 || x.indexOf("service discovery") != -1 || x.indexOf("service mesh") != -1;
  
  returnValue.hasNomad = x.indexOf("nomad") != -1 || x.indexOf("orchestration") != -1;
  
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
  
  if (stage == STAGE_D) {
    
    
     
    // To detect default meeting type sections, set meeting: to "" (empty string is --None-- in SF)
    if (MEETINGS_HAVE_DEFAULT) {
      rv = { stage : STAGE_D, meeting : "Discovery"}; 
    }
    else {
      rv = { stage : STAGE_D, meeting : ""};
    }
    
    map = discoveryMap;

  }
  else if (stage == STAGE_V) {
    
      
    // To detect default meeting type sections, set meeting: to "" (empty string is --None-- in SF)
    if (MEETINGS_HAVE_DEFAULT) {
      rv = { stage : STAGE_V, meeting : "Technical Office Hours"}; 
    }
    else {
      rv = { stage : STAGE_V, meeting : ""};
    }
    
    map = validationMap;
   
  }
  else if (stage == STAGE_C) {
    
    // Salesforce will not accept Implementation or Support in Closed/Won stage. 
    // For now we will use "--None--" to indicate Post-Sales activity!
    if (MEETINGS_HAVE_DEFAULT) {
      rv = { stage : STAGE_C, meeting : ""}; // For now this means post-sales
    }
    else {
      rv = { stage : STAGE_C, meeting : ""};
    }
    
    // Salesforce will not accept Implementation or Support in Closed/Won stage. 
    // For now we will use "--None--" to indicate Post-Sales activity!
    map = closedMap;
    
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

function filterAndAnalyzeDescription_(text) {

  let rv = {hasTeleconference : false, filteredText : text, prepTime : 0, quality : "", notes : ""}
  
  if (!text) return rv;
  
  // 
  // For the most part, assume everything after a Zoom, Webex, getclockwise, Microsoft Team Meeting
  // intros, ... to the end is garbage. (No one will add important info below all that, intentionally at least)
  //

  // Process Prep tag
  let regex = /[Pp][Rr][Ee][Pp] *: *[0-9]+[ ]?[MmHhDd]?/; // e.g. "Prep:60m" or "prep : 1H"
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
  
  // Process Notes tag
  regex = /[Nn][Oo][Tt][Ee][Ss] *: *\S+/; 
  let notesArray = text.match(regex);
  if (notesArray && notesArray[0]) {
    let i = notesArray[0].indexOf(':');
    let n = notesArray[0].substring(i+1);
    let t = n.indexOf('<');
    if (t != -1) {
      // URLs can't have a HTML tag, i.e. a < (or a > for that matter)
      n = n.substring(0,t); // prune off the tag
    }
    rv.notes = n.trim();
  }
  
  // Process Quality tag
  regex = /[Qq][Uu][Aa][Ll][Ii][Tt][Yy] *: *[0-9]/; 
  let qualityArray = text.match(regex);
  if (qualityArray && qualityArray[0]) {
    let kv = qualityArray[0].split(':');
    rv.quality = parseInt(kv[1]);
    if (rv.quality < 1) rv.quality = 1;
    else if (rv.quality > 5) rv.quality = 5;
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

