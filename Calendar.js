function createEvents(row) {
  // Get information
  var selectedClassInfo = CLASS_SCHEDULE.getActiveSheet().getRange(row, START_COLUMN_CS, 1, END_COLUMN_CS - START_COLUMN_CS + 1).getValues();
  
  /*
  // Confirmation and input which calendar is being modified
  var pass = 0;
  var cancel = 0;
  var calendarName;
  var myCalendar;
  
  while (!pass && !cancel) {
    calendarName = Browser.inputBox("Confirmation", "You are going to add events for " + selectedClassInfo[0][CLASS_ID_COLUMN - 1] + ": Level " + selectedClassInfo[0][LEVEL_COLUMN - 1] + ". \\n\\n Which Calendar do you want to add to?", Browser.Buttons.OK_CANCEL);
  
    if (calendarName == "cancel") {
      cancel = 1;
    }
    else {
      myCalendar = CalendarApp.getCalendarsByName(calendarName);
      if (myCalendar.length == 0) {
        Browser.msgBox("Warning", "The calendar doesn't exist!", Browser.Buttons.OK);
      }
      else {
        pass = 1;
      }
    }
  }
  */
  
  // Select corresponding calendar
  var calendarName = [];
  calendarName.push(getCalendarName(selectedClassInfo[0][INSTRUCTOR_COLUMN - 1]));

  // Is there a TA?
  if (selectedClassInfo[0][TA_COLUMN - 1] != "") {
    calendarName.push(getCalendarName(selectedClassInfo[0][TA_COLUMN - 1])); // There is a TA
  }
  
  // Class at BR HQ?
  var brRoom = selectedClassInfo[0][BR_ROOM_COLUMN - 1];
  if (selectedClassInfo[0][LOCATION_COLUMN - 1] == "Barnabas HQ") {
    calendarName.push("Barnabas (Company)");
  }
  
  var targetCalendar = [];
  for (var x = 0; x < calendarName.length; x++) {
    targetCalendar.push(CalendarApp.getCalendarsByName(calendarName[x]));
  }
  var calendarExist = 1;
  for (var x = 0; x < targetCalendar.length; x++) {
    if (targetCalendar[x].length == 0) {
      Browser.msgBox("Warning", "Calendar \"" + calendarName[x] + "\" doesn't exist! Please Check!", Browser.Buttons.OK);
      calendarExist = 0;
    }
  }
  
  if (calendarExist) {
    var idRecord = "";
    // Add event series  
    var title = "Level: " + selectedClassInfo[0][LEVEL_COLUMN - 1] + ", " + selectedClassInfo[0][INSTRUCTOR_COLUMN - 1];
    var startTime = new Date(selectedClassInfo[0][START_DATE_COLUMN - 1].toDateString());
    startTime.setHours(selectedClassInfo[0][START_TIME_COLUMN - 1].getHours());
    startTime.setMinutes(selectedClassInfo[0][START_TIME_COLUMN - 1].getMinutes());
    var endTime = new Date(selectedClassInfo[0][START_DATE_COLUMN - 1].toDateString());
    endTime.setHours(selectedClassInfo[0][END_TIME_COLUMN - 1].getHours());
    endTime.setMinutes(selectedClassInfo[0][END_TIME_COLUMN - 1].getMinutes());
    var endDate  = selectedClassInfo[0][END_DATE_COLUMN - 1];
    endDate.setDate(endDate.getDate() + 1);
    var location = selectedClassInfo[0][ADDRESS_COLUMN - 1];
    var note = selectedClassInfo[0][NOTE_COLUMN - 1];
    if (brRoom) { note = note + "\n" + "Barnabas HQ Room: " + brRoom; }
    
    var recurrence = CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekday(returnDay(selectedClassInfo[0][DAY_COLUMN - 1])).until(endDate);
    for (var x = 0; x < targetCalendar.length; x++) {
      var eventSeries = targetCalendar[x][0].createEventSeries(title, startTime, endTime, recurrence);
      eventSeries.setLocation(location);
      eventSeries.setDescription(note);
      
      var seriesId = eventSeries.getId();
      idRecord += calendarName[x] + "-" + seriesId + ",";
      CLASS_SCHEDULE.getActiveSheet().getRange(row, CALENDAR_ID_COLUMN).setValue(idRecord);
    }
    
    
    // Modify no class dates
    var noClassDate = selectedClassInfo[0][NO_CLASS_COLUMN - 1] + "";
    if (noClassDate != "") {
      var noClassDates = [];
      var index = noClassDate.indexOf(",");
      while (index >= 0) {
        noClassDates.push(noClassDate.substring(0, index));
        noClassDate = noClassDate.substring(index + 1).trim();
        index = noClassDate.indexOf(",");
      }
      noClassDates.push(noClassDate);
      for (var x = 0; x < noClassDates.length; x++) {
        var noClassDateStartTime = new Date(noClassDates[x]);
        noClassDateStartTime.setHours(selectedClassInfo[0][START_TIME_COLUMN - 1].getHours());
        noClassDateStartTime.setMinutes(selectedClassInfo[0][START_TIME_COLUMN - 1].getMinutes());
        var noClassDateEndTime = new Date(noClassDates[x]);
        noClassDateEndTime.setHours(selectedClassInfo[0][END_TIME_COLUMN - 1].getHours());
        noClassDateEndTime.setMinutes(selectedClassInfo[0][END_TIME_COLUMN - 1].getMinutes());
        for (var y = 0; y < targetCalendar.length; y++) {
          var noClassEvent = targetCalendar[y][0].getEvents(noClassDateStartTime, noClassDateEndTime, {search: title});
          if (noClassEvent[0]) {
            noClassEvent[0].setTitle("[NO CLASS] " + noClassEvent[0].getTitle());
          }
          else {
            var d = new Date(noClassDates[x]);
            Browser.msgBox("Warning", "No class date: " + d.toLocaleDateString() + " is invalid! Position: Row " + row + 
                           "\\n \\n Possible mistakes:" + 
                           "\\n 1. This date is not a " +  selectedClassInfo[0][DAY_COLUMN - 1] + "." +
                           "\\n 2. This date is beyond the class timeline." + 
                           "\\n \\n This date will be skipped for now, please modify and update this class again after this updating.", Browser.Buttons.OK);
          }
        }
      }
    }
    
    // Add additional class dates
    var additionalDate = selectedClassInfo[0][ADDITIONAL_CLASS_COLUMN - 1] + "";
    if (additionalDate != "") {
      var additionalDates = [];
      var index = additionalDate.indexOf(",");
      while (index >= 0) {
        additionalDates.push(additionalDate.substring(0,index));
        additionalDate = additionalDate.substring(index + 1).trim();
        index = additionalDate.indexOf(",");
      }
      additionalDates.push(additionalDate);
      for (var x = 0; x < additionalDates.length; x++) {
        var additionalDateStartTime = new Date(additionalDates[x]);
        additionalDateStartTime.setHours(selectedClassInfo[0][START_TIME_COLUMN - 1].getHours());
        additionalDateStartTime.setMinutes(selectedClassInfo[0][START_TIME_COLUMN - 1].getMinutes());
        var additionalDateEndTime = new Date(additionalDates[x]);
        additionalDateEndTime.setHours(selectedClassInfo[0][END_TIME_COLUMN - 1].getHours());
        additionalDateEndTime.setMinutes(selectedClassInfo[0][END_TIME_COLUMN - 1].getMinutes());
        for (var y = 0; y < targetCalendar.length; y++) {
          var additionalEvent = targetCalendar[y][0].createEvent(title, additionalDateStartTime, additionalDateEndTime);
          additionalEvent.setLocation(location);
          additionalEvent.setDescription(note);
          var eventId = additionalEvent.getId();
          idRecord += calendarName[y] + "-" + eventId + ",";
          CLASS_SCHEDULE.getActiveSheet().getRange(row, CALENDAR_ID_COLUMN).setValue(idRecord);
        }
      }
    }
    
    // Log
    var operation = "Update/Add to Calendar";
    var classId = selectedClassInfo[0][CLASS_ID_COLUMN - 1];
    var calendarString = "";
    for (var i = 0; i < calendarName.length; i++) {
      calendarString = calendarString + calendarName[i];
      if (i != calendarName.length - 1) {
        calendarString += ", ";
      }
    }
    var description = "Class: " + classId + " has been updated/added to calendar(s): " + calendarString + ".";
    log(operation, description, "");
  }
}

/*
function setHolidays() {
  var another = 1;
  var cancel = 0;
  while (another && !cancel) {
    var holidayDate = Browser.inputBox("Set Holiday", "Please enter the holiday date(mm/dd/yyyy): ", Browser.Buttons.OK_CANCEL);
    if (holidayDate == "cancel") {
      cancel = 1;
    }
    else {
      var noClassDate = holidayDate;
      var noClassDateStartTime = new Date(noClassDate);
      if (noClassDateStartTime == "Invalid Date") {
        Browser.msgBox("The date is invalid!");
      }
      else {
        var selectedCell = classSchedule.getActiveSheet().getActiveCell();
        var selectedRow = selectedCell.getRow();
        var selectedClassInfo = classSchedule.getActiveSheet().getRange(selectedRow, START_COLUMN_CS, 1, END_COLUMN_CS - START_COLUMN_CS + 1).getValues();
        noClassDateStartTime.setHours(selectedClassInfo[0][START_TIME_COLUMN - 1].getHours());
        noClassDateStartTime.setMinutes(selectedClassInfo[0][START_TIME_COLUMN - 1].getMinutes());
        var noClassDateEndTime = new Date(noClassDate);
        noClassDateEndTime.setHours(selectedClassInfo[0][END_TIME_COLUMN - 1].getHours());
        noClassDateEndTime.setMinutes(selectedClassInfo[0][END_TIME_COLUMN - 1].getMinutes());
        
        var myCalendar = CalendarApp.getCalendarsByName("Test"); // Remember to change the calendar name.
        var noClassEvent = myCalendar[0].getEvents(noClassDateStartTime, noClassDateEndTime);
        noClassEvent[0].deleteEvent();
        var moreHoliday = Browser.msgBox("More Holiday?", "Do you want to add another holiday?", Browser.Buttons.YES_NO);
        if (moreHoliday == "yes") {
          another = 1;
        }
        else {
          another = 0;
        }
      }
    }
  }
}
*/

function deleteEvents(row) {
  var selectedClassInfo = CLASS_SCHEDULE.getActiveSheet().getRange(row, START_COLUMN_CS, 1, END_COLUMN_CS - START_COLUMN_CS + 1).getValues();
  
  var idRecord = selectedClassInfo[0][CALENDAR_ID_COLUMN - 1];
  if (idRecord == "") {
    Browser.msgBox("Warning", "Class: " + selectedClassInfo[0][CLASS_ID_COLUMN - 1] + " has never been added to a calendar!", Browser.Buttons.OK);
  }// no events exist
  else {
    var calendars = [];
    var eventIds = [];
    while (idRecord.indexOf(",") >= 0) {
      calendars.push(idRecord.substring(0, idRecord.indexOf("-")));
      idRecord = idRecord.substring(idRecord.indexOf("-") + 1);
      eventIds.push(idRecord.substring(0, idRecord.indexOf(",")));
      idRecord = idRecord.substring(idRecord.indexOf(",") + 1);
    }
    for (var x = 0; x < calendars.length; x++) {
      var targetCalendar = CalendarApp.getCalendarsByName(calendars[x]);
      if (targetCalendar[0].getEventSeriesById(eventIds[x])) {
        try { targetCalendar[0].getEventSeriesById(eventIds[x]).deleteEventSeries(); }
        catch (error) {}
      }
      else if (targetCalendar[0].getEventById(eventIds[x])) {
        try { targetCalendar[0].getEventById(eventIds[x]).deleteEvent(); }
        catch (error) {}
      }
    }
    CLASS_SCHEDULE.getActiveSheet().getRange(row, CALENDAR_ID_COLUMN).clearContent();
    if (calendars.length > 0) {
      // Log
      var operation = "Delete from Calendar";
      var classId = selectedClassInfo[0][CLASS_ID_COLUMN - 1];
      var calendarString = "";
      for (var i = 0; i < calendars.length; i++) {
        calendarString = calendarString + calendars[i];
        if (i != calendars.length - 1) {
          calendarString += ", ";
        }
      }
      var description = "Class: " + classId + " has been removed from calendar(s): " + calendarString + ".";
      log(operation, description, "");
    }
  }
}

  /*
  var seriesId = CLASS_SCHEDULE.getActiveSheet().getRange(selectedRow, CALENDAR_ID_COLUMN).getValue();
  var calendarName = "Barnabas (" + getFirstName(selectedClassInfo[0][INSTRUCTOR_COLUMN - 1]) + ")";
  var myCalendar = CalendarApp.getCalendarsByName(calendarName);
  myCalendar[0].getEventSeriesById(seriesId).deleteEventSeries();
  CLASS_SCHEDULE.getActiveSheet().getRange(selectedRow, CALENDAR_ID_COLUMN).clearContent();
  
  // Delete additional events
  var additionalDate = selectedClassInfo[0][ADDITIONAL_CLASS_COLUMN - 1] + "";
  if (additionalDate != "") {
    var additionalDates = [];
    var index = additionalDate.indexOf(",");
    while (index >= 0) {
      additionalDates.push(additionalDate.substring(0,index));
      additionalDate = additionalDate.substring(index + 1).trim();
      index = additionalDate.indexOf(",");
    }
    additionalDates.push(additionalDate);
    for (var x = 0; x < additionalDates.length; x++) {
      var additionalDateStartTime = new Date(additionalDates[x]);
      additionalDateStartTime.setHours(selectedClassInfo[0][START_TIME_COLUMN - 1].getHours());
      additionalDateStartTime.setMinutes(selectedClassInfo[0][START_TIME_COLUMN - 1].getMinutes());
      var additionalDateEndTime = new Date(additionalDates[x]);
      additionalDateEndTime.setHours(selectedClassInfo[0][END_TIME_COLUMN - 1].getHours());
      additionalDateEndTime.setMinutes(selectedClassInfo[0][END_TIME_COLUMN - 1].getMinutes());
      var additionalEvent = myCalendar[0].getEvents(additionalDateStartTime, additionalDateEndTime);
      additionalEvent[0].deleteEvent();
    }
  }
  */

function updateEvents() {
  // get selected rows
  var activeSheet = CLASS_SCHEDULE.getActiveSheet();
  var lastRow = activeSheet.getLastRow();
  if (lastRow < START_ROW_CS) { return; }
  var selectStatus = activeSheet.getRange(START_ROW_CS, SELECT_COLUMN, lastRow - START_ROW_CS + 1, 1).getValues();
  var selectedRows = [];
  for (var x = 0; x < selectStatus.length; x++) {
    if (selectStatus[x][0] == true) {
      selectedRows.push(START_ROW_CS + x);
    }
  }
  // if any row is selected
  if (selectedRows.length > 0  && calendarsExist(selectedRows)) {
    // update for each row
    for (var x = 0; x < selectedRows.length; x++) {
      if (eventsExist(selectedRows[x])) {
        deleteEvents(selectedRows[x]);
      }
      createEvents(selectedRows[x]);
    }
    Browser.msgBox("Success", "Calendars have been updated!", Browser.Buttons.OK);
  }
}

function deleteClasses() {
  var confirmation = Browser.msgBox("Confirmation", "Are you sure you want to remove this class from the corresponding calendar?", Browser.Buttons.YES_NO);
  if (confirmation == "no") { return; }
  // get selected rows
  var activeSheet = CLASS_SCHEDULE.getActiveSheet();
  var selectStatus = activeSheet.getRange(START_ROW_CS, SELECT_COLUMN, activeSheet.getLastRow() - START_ROW_CS + 1, 1).getValues();
  var selectedRows = [];
  var selectedClasses = [];
  var calendarArr = [];
  for (var x = 0; x < selectStatus.length; x++) {
    if (selectStatus[x][0] == true) {
      selectedRows.push(START_ROW_CS + x);
      selectedClasses.push(activeSheet.getRange(START_ROW_CS + x, CLASS_ID_COLUMN).getValue());
      var calendars = getCalendarsForRow(START_ROW_CS + x);
      calendarArr.push(calendars);
    }
  }
  // if any row is selected
  if (selectedRows.length > 0) {
    // delete events for each row
    for (var x = 0; x < selectedRows.length; x++) {
      if (eventsExist(selectedRows[x])) {
        deleteEvents(selectedRows[x]);
      }
    }
    Browser.msgBox("Success", "Calendars have been deleted!", Browser.Buttons.OK);
    if (isAutoEmailEnabled) {
      var classNo = selectedClasses.length;
      var subject = classNo + " class(es) has/have been removed from calendar(s).";
      var text = "";
      for (var i = 0; i < classNo; i++) {
        var classId = selectedClasses[i];
        text += "Class " + classId + " has been removed from calendars: ";
        for (var j = 0; j < calendarArr[i].length; j++) {
          var calendar = calendarArr[i][j];
          text += calendar;
          if (j != calendarArr[i].length - 1) {
            text += ", ";
          }
        }
        text += ".\n";
      }
      autoEmail(recipient, subject, text);
    }
  }
}

function eventsExist(row) {
  if (CLASS_SCHEDULE.getActiveSheet().getRange(row, CALENDAR_ID_COLUMN).getValue()) { return 1; }
  else { return 0; }
}

function returnDay(weekDay) {
  if(weekDay == "Monday") {
    return CalendarApp.Weekday.MONDAY;
  } 
  else if(weekDay == "Tuesday") {
    return CalendarApp.Weekday.TUESDAY;
  } 
  else if(weekDay == "Wednesday") {
    return CalendarApp.Weekday.WEDNESDAY;
  }
  else if(weekDay == "Thursday") {
    return CalendarApp.Weekday.THURSDAY;
  }
  else if(weekDay == "Friday") {
    return CalendarApp.Weekday.FRIDAY;
  }
  else if(weekDay == "Saturday") {
    return CalendarApp.Weekday.SATURDAY;
  }
  else {
    return CalendarApp.Weekday.SUNDAY;
  }
}

function calendarsExist(rows) {
  for (var x = 0; x < rows.length; x++) {
    var instructors = [];
    instructors.push(CLASS_SCHEDULE.getActiveSheet().getRange(rows[x], INSTRUCTOR_COLUMN).getValue());
    if (CLASS_SCHEDULE.getActiveSheet().getRange(rows[x], TA_COLUMN).getValue()) {
      instructors.push(CLASS_SCHEDULE.getActiveSheet().getRange(rows[x], TA_COLUMN).getValue());
    }
    for (var y = 0; y < instructors.length; y++) {
      var calendarName = getCalendarName(instructors[y]);
      if (!calendarName) {
        Browser.msgBox("Warning", instructors[y] + " doesn't have a calendar!", Browser.Buttons.OK);
        return 0;
      }
      else {
        var calendar = CalendarApp.getCalendarsByName(calendarName);
        if (calendar.length == 0) {
          Browser.msgBox("Warning", "Calendar \"" + calendarName + "\" doesn't exist!", Browser.Buttons.OK);
          return 0;
        }
      }
    }
  }
  return 1;
}