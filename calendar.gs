function import_calendars() {

  clearTab_(CALENDAR, CALENDAR_HEADER);
  clearTab_(LOG_TAB);
  
  let sheet = SpreadsheetApp.getActive().getSheetByName(CALENDAR);
  let importCursor = {sheet : sheet, row : 2}; // Drop to row below the header we kept during the clear
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
  
  sheet = SpreadsheetApp.getActive().getSheetByName(LOG_TAB);
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


