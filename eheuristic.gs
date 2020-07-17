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
    hasNomad : false,
    count : function () {
      let retval = 0;
      if (this.hasTerraform) retval++;
      if (this.hasVault) retval++;
      if (this.hasNomad) retval++;
      if (this.hasConsul) retval++;
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
  
  returnValue.hasVault = x.indexOf("vault") != -1 || x.indexOf("secrets management") != -1;
  
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

function filterAndAnalyzeDescription_(text) {

  let rv = {hasTeleconference : false, filteredText : text, prepTime : 0, quality : ""}

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
  
  // Process Quality tag
  regex = /[Qq][Uu][Au][Ll][Ii][Tt][Yy] *: *[0-9]/; 
  let qualArray = text.match(regex);
  if (qualArray && qualArray[0]) {
    let kv = qualArray[0].split(':');
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

