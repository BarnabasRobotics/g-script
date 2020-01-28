function getClassLevel() {
  var classLevels = [];
  var row = START_ROW_RS;
  while (RESOURCE_SHEET.getRange(row, LEVEL_COLUMN_RS).getValue()) {
    classLevels.push(RESOURCE_SHEET.getRange(row, LEVEL_COLUMN_RS).getValue());
    row++;
  }
  return classLevels;
}

function getInstructors() {
  var instructors = [];
  var row = START_ROW_RS;
  while (RESOURCE_SHEET.getRange(row, INSTRUCTOR_COLUMN_RS).getValue()) {
    instructors.push(RESOURCE_SHEET.getRange(row, INSTRUCTOR_COLUMN_RS).getValue());
    row++;
  }
  return instructors;
}

function getStatus() {
  var Statuses = [];
  var row = START_ROW_RS;
  while (RESOURCE_SHEET.getRange(row, STATUS_COLUMN_RS).getValue()) {
    Statuses.push(RESOURCE_SHEET.getRange(row, STATUS_COLUMN_RS).getValue());
    row++;
  }
  return Statuses;
}

function getAccountTypes() {
  var accountTypes = [];
  var row = START_ROW_RS;
  while (RESOURCE_SHEET.getRange(row, ACCOUNT_TYPE_COLUMN_RS).getValue()) {
    accountTypes.push(RESOURCE_SHEET.getRange(row, ACCOUNT_TYPE_COLUMN_RS).getValue());
    row++;
  }
  return accountTypes;
}

function getFirstName(name) {
  return name.substring(0, name.indexOf(" "));
}

function getCalendarName(instructor) {
  var instructors = getInstructors();
  var index = instructors.indexOf(instructor);
  if (index == -1) {
    Browser.msgBox("Error", "Can't find \"" + instructor + "\"", Browser.Buttons.OK);
    return 0;
  }
  else {
    return RESOURCE_SHEET.getRange(START_ROW_RS + index, CALENDAR_NAME_COLUMN_RS).getValue();
  }
}

function getMonths(startDate, endDate) {
  var months = [];
  for (var x = new Date(startDate.getYear(), startDate.getMonth()); x <= endDate; x.setMonth(x.getMonth() + 1)) {
    var month = MONTHS[x.getMonth()];
    months.push(month);
  }
  return months;
}

function getMonthsIndex(startDate, endDate) {
  var months = [];
  for (var x = new Date(startDate.getYear(), startDate.getMonth()); x <= endDate; x.setMonth(x.getMonth() + 1)) {
    var month = x.getMonth();
    months.push(month);
  }
  return months;
}

function durationToSecond(duration) {
  var hours = duration.getHours();
  var minutes = duration.getMinutes();
  var seconds = duration.getSeconds();
  return hours * 3600 + minutes * 60 + seconds;
}

function secondToDurationText(seconds) {
  var hours = Math.floor(seconds / 3600);
  seconds = seconds - hours * 3600;
  var minutes = Math.floor(seconds / 60);
  seconds = seconds - minutes * 60;
  return hours + ":" + minutes + ":" + seconds;
}

function getClassDataRange(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow >= START_ROW_CS) {
    var classDataRange = sheet.getRange(START_ROW_CS, START_COLUMN_CS, lastRow - START_ROW_CS + 1, sheet.getMaxColumns());
    return classDataRange;
  }
  else {
    return 0;
  }
}

function removeInvalidDates(dates) {
  for (var x = 0; x < dates.length; x++) {
    if (dates[x][0] == "TBD" || !dates[x][0]) {
      dates.splice(x, 1);
      x--;
    }
  }
  return dates;
}

function getCalendarsForRow(row) {
  var selectedClassInfo = CLASS_SCHEDULE.getActiveSheet().getRange(row, START_COLUMN_CS, 1, END_COLUMN_CS - START_COLUMN_CS + 1).getValues();
  
  // Select corresponding calendar
  var calendars = [];
  calendars.push(getCalendarName(selectedClassInfo[0][INSTRUCTOR_COLUMN - 1]));

  // Is there a TA?
  if (selectedClassInfo[0][TA_COLUMN - 1] != "") {
    calendars.push(getCalendarName(selectedClassInfo[0][TA_COLUMN - 1])); // There is a TA
  }
  
  // Class at BR HQ?
  if (selectedClassInfo[0][LOCATION_COLUMN - 1] == "Barnabas HQ") {
    calendars.push("Barnabas (Company)");
  }
  
  return calendars;
}