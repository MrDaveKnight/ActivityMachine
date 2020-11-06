function import_calendars() {

  clearTab_(CALENDAR, CALENDAR_HEADER);
  clearTab_(LOG_TAB);
  
  logStamp("Calendar Import");
  
  let sheet = SpreadsheetApp.getActive().getSheetByName(CALENDAR);
  let importCursor = {sheet : sheet, row : 2}; // Drop to row below the header we kept during the clear
  let parms = SpreadsheetApp.getActive().getSheetByName(RUN_PARMS);
  
  let dateRange = parms.getRange(3,2,2,1); // Hardcoded to format in RUN_PARMS
  let dates = dateRange.getValues();
  
  let calendarRange = parms.getRange(3,5,25,1); // Hardcoded to format in RUN_PARMS
  let calendars = calendarRange.getValues();
  
  let inviteCount = 0;
  for (var row in calendars) {
    if (calendars[row][0]) {      
      inviteCount += import_gcal_(calendars[row][0], dates[0][0], dates[1][0], importCursor);
    }
  }
  logOneCol("Imported a total of " + inviteCount + " invites.");
  logOneCol("End time: " + new Date().toLocaleTimeString());
  
  createDataLoadFilters(); // Just setup for event processing
  createChoiceLists();
  logOneCol("End time: " + new Date().toLocaleTimeString());
  
}

function import_gcal_(calName, startDate, endDate, cursor){
  
  //
  // Export Google Calendar Events to a Google Spreadsheet
  //
  // This function retrieves events between 2 dates for the specified calendar.
  // It saves the eveents in the Calendar tab of the current spreadsheet starting at the second row (after the header).
  // Returns the number of invites imported.
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
  
  // Loop through all calendar events found and write them out to the Calendar tab
  for (var i=0;i<events.length;i++) {
  
    // Collect some quick stats
    var isCustEvent = false;
    var attendees=events[i].getGuestList(true);
    var attendeeList = "";
    if (attendees.length > 0) {
      attendeeList = attendees[0].getEmail();
    }
    for (var j=1; j < attendees.length; j++) {
      attendeeList = attendeeList + "," + attendees[j].getEmail();
    }
      
    // No tags show up for Notes
    /*
    let tags = events[i].getAllTagKeys();
    let tagNames = "";
    if (tags && tags.length > 0) {
      tagNames = tags[0];
      for (let x = 1; x < tags.length; x++) {
        tagNames += ", " + tags[x];
      }
       Logger.log("Tag:" + tagNames);
    }
    */
    
    
    // The details below must match the header defined in main.gs
    let details=[[calName, events[i].getTitle(), events[i].getLocation(), events[i].getStartTime(), events[i].getEndTime(), events[i].getMyStatus(), events[i].getCreators(), events[i].isAllDayEvent(), events[i].isRecurringEvent(), attendeeList, events[i].getDescription()]];
    let range=cursor.sheet.getRange(cursor.row,1,1,11); // 11 must match details length
    range.setValues(details);
    cursor.row++;
  }
  
  logOneCol("Imported " + calName + ", " + events.length + " invites.");
  
  return events.length;
}


