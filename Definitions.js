// Definitions

// This spreadsheet
var CLASS_SCHEDULE = SpreadsheetApp.getActiveSpreadsheet();
var RESOURCE_SHEET = CLASS_SCHEDULE.getSheetByName("Resources");
var TEMPLATE_SHEET = CLASS_SCHEDULE.getSheetByName("Class Schedule Template");
var LOG_SHEET = CLASS_SCHEDULE.getSheetByName("Log");

// Definitions of all class sheets - CS
var HEAD_ROW_CS = 2;
var START_ROW_CS = 3;
var START_COLUMN_CS = 1;
var END_COLUMN_CS = 25;
var SELECT_COLUMN = 1;
var CLASS_ID_COLUMN = 2;
var START_DATE_COLUMN = 3;
var END_DATE_COLUMN = 4;
var DAY_COLUMN = 5;
var LEVEL_COLUMN = 6;
var START_TIME_COLUMN = 7;
var END_TIME_COLUMN = 8;
var LOCATION_COLUMN = 9;
var BR_ROOM_COLUMN = 10;
var INSTRUCTOR_COLUMN = 11;
var STATUS_COLUMN = 12;
var ADDRESS_COLUMN = 13;
var NO_CLASS_COLUMN = 14;
var ADDITIONAL_CLASS_COLUMN = 15;
var NOTE_COLUMN = 16;
var ACCOUNT_TYPE_COLUMN = 17;
var TA_COLUMN = 18;
var STUDENT_NUM_COLUMM = 19;
var WEEK_HR_COLUMN = 22;
var CALENDAR_ID_COLUMN = 23;
var DAY_INDEX_COLUMN = 24;
var STATUS_INDEX_COLUMN = 25;

// Definitions of resource sheet - RS
var HEAD_ROW_RS = 1;
var START_ROW_RS = 2;
var LEVEL_COLUMN_RS = 6;
var INSTRUCTOR_COLUMN_RS = 2;
var CALENDAR_NAME_COLUMN_RS = 3;
var STATUS_COLUMN_RS = 4;
var ACCOUNT_TYPE_COLUMN_RS = 9;

// Definitions of log sheet - LG
var HEAD_ROW_LG = 1;
var START_ROW_LG = 2;
var DATE_COLUMN_LG = 1;
var TIME_COLUMN_LG = 2;
var OPERATION_COLUMN_LG = 3;
var DESC_COLUMN_LG = 4;
var USER_COLUMN_LG = 5;
var NOTE_COLUMN_LG = 6;

// Definitions of reports sheet
var EMPTY_LINE_WIDTH = 20;

// Resource variables
var DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
var MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

// Auto email feature
var isAutoEmailEnabled = true;
var recipient = "class@barnabasrobotics.com";
var self = "ling@barnabasrobotics.com";

// autoEmail function
function autoEmail(recipient, subject, body) {
  var options = {cc: self};
  GmailApp.sendEmail(recipient, subject, body, options);
}

// Log operations
function log(operation, description, note) {
  LOG_SHEET.insertRowBefore(START_ROW_LG);
  // define log info range
  var dateCell = LOG_SHEET.getRange(START_ROW_LG, DATE_COLUMN_LG); 
  var timeCell = LOG_SHEET.getRange(START_ROW_LG, TIME_COLUMN_LG);  
  var operationCell = LOG_SHEET.getRange(START_ROW_LG, OPERATION_COLUMN_LG); 
  var descCell = LOG_SHEET.getRange(START_ROW_LG, DESC_COLUMN_LG);
  var userCell = LOG_SHEET.getRange(START_ROW_LG, USER_COLUMN_LG);  
  var noteCell = LOG_SHEET.getRange(START_ROW_LG, NOTE_COLUMN_LG); 
  // set date and time 
  var d = new Date();
  var logDate = d.toDateString();
  var logTimeLong = d.toTimeString();
  var logTime = logTimeLong.substring(0, logTimeLong.search(" "));
  dateCell.setValue(logDate);
  timeCell.setValue(logTime);
  // set operation
  operationCell.setValue(operation);
  // set description
  descCell.setValue(description);
  // set user
  var logUser = Session.getActiveUser().getEmail();
  userCell.setValue(logUser);
  // set note
  noteCell.setValue(note); 
}

// update reports
function updateReports() {
  // define target sheet
  var classSheet = CLASS_SCHEDULE.getActiveSheet();
  var reportSheetName = classSheet.getName() + " Reports";
  var reportSheet;
  if (reportSheet = CLASS_SCHEDULE.getSheetByName(reportSheetName)) {
    CLASS_SCHEDULE.deleteSheet(reportSheet);
  }
  CLASS_SCHEDULE.insertSheet(reportSheetName);
  reportSheet = CLASS_SCHEDULE.getSheetByName(reportSheetName);
  
  // get information from class schedule sheet
  var startDates = classSheet.getRange(START_ROW_CS, START_DATE_COLUMN, classSheet.getLastRow() - START_ROW_CS + 1, 1).getValues();
  startDates = removeInvalidDates(startDates);
  var endDates = classSheet.getRange(START_ROW_CS, END_DATE_COLUMN, classSheet.getLastRow() - START_ROW_CS + 1, 1).getValues();
  endDates = removeInvalidDates(endDates);
  var instructors = classSheet.getRange(START_ROW_CS, INSTRUCTOR_COLUMN, classSheet.getLastRow() - START_ROW_CS + 1, 1).getValues();
  var TAs = classSheet.getRange(START_ROW_CS, TA_COLUMN, classSheet.getLastRow() - START_ROW_CS + 1, 1).getValues();
  var classes = classSheet.getRange(START_ROW_CS, LEVEL_COLUMN, classSheet.getLastRow() - START_ROW_CS + 1, 1).getValues();
  var earliestDate = startDates[0][0];
  var latestDate = endDates[0][0];
  for (var x = 1; x < startDates.length; x++) {
    if (startDates[x][0] < earliestDate) {
      earliestDate = startDates[x][0];
    }
  }
  for (var x = 1; x < endDates.length; x++) {
    if (endDates[x][0] > latestDate) {
      latestDate = endDates[x][0];
    }
  }
  var reportMonths = getMonths(earliestDate, latestDate);
  var reportMonthsIndex = getMonthsIndex(earliestDate, latestDate);
  var reportInstructors = [];
  for (var x = 0; x < instructors.length; x++) {
    if (reportInstructors.indexOf(getFirstName(instructors[x][0])) == -1) {
      reportInstructors.push(getFirstName(instructors[x][0]));    
    }
    if (TAs[x][0] && reportInstructors.indexOf(getFirstName(TAs[x][0])) == -1) {
      reportInstructors.push(getFirstName(TAs[x][0]));
    }
  }
  var reportClasses = [];
  for (var x = 0; x < classes.length; x++) {
    if (reportClasses.indexOf(classes[x][0]) == -1) {
      reportClasses.push(classes[x][0]);
    }
  }
  
  // position of different reports
  // monthly teaching hours (confirmed) - mthc
  var mthcFirstRow = 1;
  var mthcRows = 3 + reportMonths.length;
  var mthcFirstColumn = 1;
  var mthcColumns = 2 + reportInstructors.length;
  var mthcInstructorRow = 2;
  var mthcMonthColumn = 1;
  // division column 1
  var dColumn1 = mthcFirstColumn + mthcColumns;
  reportSheet.setColumnWidth(dColumn1, EMPTY_LINE_WIDTH);
  // monthly teaching hours (pending) = mthp
  var mthpFirstRow = 1;
  var mthpRows = 3 + reportMonths.length;
  var mthpFirstColumn = dColumn1 + 1;
  var mthpColumns = 2 + reportInstructors.length;
  var mthpInstructorRow = 2;
  var mthpMonthColumn = mthpFirstColumn;
  // division row 1
  var dRow1 = mthcFirstRow + mthcRows;
  reportSheet.setRowHeight(dRow1, EMPTY_LINE_WIDTH);
  // monthly teaching hours (total) - mtht
  var mthtFirstRow = dRow1 + 1;
  var mthtRows = 3 + reportMonths.length;
  var mthtFirstColumn = 1;
  var mthtColumns = 2 + reportInstructors.length;
  var mthtInstructorRow = mthtFirstRow + 1;
  var mthtMonthColumn = 1;
  // division row 2
  var dRow2 = mthtFirstRow + mthtRows;
  reportSheet.setRowHeight(dRow2, EMPTY_LINE_WIDTH);
  // average weekly teaching hours (confirmed) - wthc
  var wthcFirstRow = dRow2 + 1;
  var wthcRows = 3 + reportMonths.length;
  var wthcFirstColumn = 1;
  var wthcColumns = 2 + reportInstructors.length;
  var wthcInstructorRow = wthcFirstRow + 1;
  var wthcMonthColumn = 1;
  // average weekly teaching hours (pending) - wthp
  var wthpFirstRow = dRow2 + 1;
  var wthpRows = 3 + reportMonths.length;
  var wthpFirstColumn = dColumn1 + 1;
  var wthpColumns = 2 + reportInstructors.length;
  var wthpInstructorRow = wthpFirstRow + 1;
  var wthpMonthColumn = mthpFirstColumn;
  // division row 3
  var dRow3 = wthcFirstRow + wthcRows;
  // average weekly teaching hours (total) - wtht
  var wthtFirstRow = dRow3 + 1;
  var wthtRows = 3 + reportMonths.length;
  var wthtFirstColumn = 1;
  var wthtColumns = 2 + reportInstructors.length;
  var wthtInstructorRow = wthtFirstRow + 1;
  var wthtMonthColumn = wthtFirstColumn;
  // division row 3
  var dRow4 = wthtFirstRow + wthtRows;
  // attendence report  - att
  var attFirstRow = dRow4 + 1;
  var attRows = 3 + reportClasses.length;
  var attFirstColumn = 1;
  var attColumns = 2;
  var attHeadRow = attFirstRow + 1;
  var attClassColumn = attFirstColumn;
  var attStudentNumColumn = attClassColumn + 1;
  // set first column width
  reportSheet.setColumnWidth(mthcFirstColumn, 150);
  reportSheet.setColumnWidth(mthpFirstColumn, 150);
  
  // file title row
  // monthly teaching hours (confirmed) - mthc
  reportSheet.getRange(mthcFirstRow, mthcFirstColumn).setValue("Monthly Teaching Hours (Confirmed):");
  reportSheet.getRange(mthcFirstRow, mthcFirstColumn, 1, mthcColumns).merge();
  reportSheet.getRange(mthcInstructorRow, mthcFirstColumn).setValue("Month \\ Instructor");
  // monthly teaching hours (pending) - mthp
  reportSheet.getRange(mthpFirstRow, mthpFirstColumn).setValue("Monthly Teaching Hours (Pending):");
  reportSheet.getRange(mthpFirstRow, mthpFirstColumn, 1, mthpColumns).merge();
  reportSheet.getRange(mthpInstructorRow, mthpFirstColumn).setValue("Month \\ Instructor");
  // monthly teaching hours (total) - mtht
  reportSheet.getRange(mthtFirstRow, mthtFirstColumn).setValue("Monthly Teaching Hours (Total):");
  reportSheet.getRange(mthtFirstRow, mthtFirstColumn, 1, mthtColumns).merge();
  reportSheet.getRange(mthtInstructorRow, mthtFirstColumn).setValue("Month \\ Instructor");
  // average weekly teaching hours (confirmed) - wthc
  reportSheet.getRange(wthcFirstRow, wthcFirstColumn).setValue("Average Weekly Teaching Hours (Confirmed):");
  reportSheet.getRange(wthcFirstRow, wthcFirstColumn, 1, wthcColumns).merge();
  reportSheet.getRange(wthcInstructorRow, wthcFirstColumn).setValue("Month \\ Instructor");
  // average weekly teaching hours (pending) - wthp
  reportSheet.getRange(wthpFirstRow, wthpFirstColumn).setValue("Average Weekly Teaching Hours (Pending):");
  reportSheet.getRange(wthpFirstRow, wthpFirstColumn, 1, wthpColumns).merge();
  reportSheet.getRange(wthpInstructorRow, wthpFirstColumn).setValue("Month \\ Instructor");
  // average weekly teaching hours (total) - wtht
  reportSheet.getRange(wthtFirstRow, wthtFirstColumn).setValue("Average Weekly Teaching Hours (Total):");
  reportSheet.getRange(wthtFirstRow, wthtFirstColumn, 1, wthtColumns).merge();
  reportSheet.getRange(wthtInstructorRow, wthtFirstColumn).setValue("Month \\ Instructor");
  // attendence report - att
  reportSheet.getRange(attFirstRow, attFirstColumn).setValue("Attendence (Comfirmed)");
  reportSheet.getRange(attFirstRow, attFirstColumn, 1, attColumns).merge();
  reportSheet.getRange(attHeadRow, attFirstColumn).setValue("Class");
  reportSheet.getRange(attHeadRow, attStudentNumColumn).setValue("Student Num");
  reportSheet.getRange(attFirstRow + attRows - 1, attFirstColumn).setValue("Total");
  for (var x = 0; x < reportClasses.length; x++) {
    reportSheet.getRange(attHeadRow + 1 + x, attClassColumn).setValue(reportClasses[x]);
  }
  
  // Set instructor row
  for (var x = 0; x < reportInstructors.length; x++) {
    // monthly teaching hours (confirmed) - mthc
    reportSheet.getRange(mthcInstructorRow, mthcFirstColumn + 1 + x).setValue(reportInstructors[x]);
    // monthly teaching hours (pending) - mthp
    reportSheet.getRange(mthpInstructorRow, mthpFirstColumn + 1 + x).setValue(reportInstructors[x]);
    // monthly teaching hours (total) - mtht
    reportSheet.getRange(mthtInstructorRow, mthtFirstColumn + 1 + x).setValue(reportInstructors[x]);
    // average weekly teaching hours (confirmed) - wthc
    reportSheet.getRange(wthcInstructorRow, wthcFirstColumn + 1 + x).setValue(reportInstructors[x]);
    // average weekly teaching hours (pending) - wthp
    reportSheet.getRange(wthpInstructorRow, wthpFirstColumn + 1 + x).setValue(reportInstructors[x]);
    // average weekly teaching hours (total) - wtht
    reportSheet.getRange(wthtInstructorRow, wthtFirstColumn + 1 + x).setValue(reportInstructors[x]);
  }
  // monthly teaching hours (confirmed) - mthc
  reportSheet.getRange(mthcInstructorRow, mthcFirstColumn + mthcColumns - 1).setValue("Total");
  // monthly teaching hours (pending) - mthp
  reportSheet.getRange(mthpInstructorRow, mthpFirstColumn + mthpColumns - 1).setValue("Total");
  // monthly teaching hours (total) - mtht
  reportSheet.getRange(mthtInstructorRow, mthtFirstColumn + mthtColumns - 1).setValue("Total");
  // average weekly teaching hours (confirmed) - wthc
  reportSheet.getRange(wthcInstructorRow, wthcFirstColumn + wthcColumns - 1).setValue("Total");
  // average weekly teaching hours (pending) - wthp
  reportSheet.getRange(wthpInstructorRow, wthpFirstColumn + wthpColumns - 1).setValue("Total");
  // average weekly teaching hours (total) - wtht
  reportSheet.getRange(wthtInstructorRow, wthtFirstColumn + wthtColumns - 1).setValue("Total");
  
  // set month column
  for (var x = 0; x < reportMonths.length; x++) {
    // monthly teaching hours (confirmed) - mthc
    reportSheet.getRange(mthcInstructorRow + 1 + x, mthcMonthColumn).setValue(reportMonths[x]);
    // monthly teaching hours (pending) - mthp
    reportSheet.getRange(mthpInstructorRow + 1 + x, mthpMonthColumn).setValue(reportMonths[x]);
    // monthly teaching hours (total) - mtht
    reportSheet.getRange(mthtInstructorRow + 1 + x, mthtMonthColumn).setValue(reportMonths[x]);
    // average weekly teaching hours (confirmed) - wthc
    reportSheet.getRange(wthcInstructorRow + 1 + x, wthcMonthColumn).setValue(reportMonths[x]);
    // average weekly teaching hours (pending) - wthp
    reportSheet.getRange(wthpInstructorRow + 1 + x, wthpMonthColumn).setValue(reportMonths[x]);
    // average weekly teaching hours (total) - wtht
    reportSheet.getRange(wthtInstructorRow + 1 + x, wthtMonthColumn).setValue(reportMonths[x]);
  }
  // monthly teaching hours (confirmed) - mthc
  reportSheet.getRange(mthcFirstRow + mthcRows - 1, mthcMonthColumn).setValue("Total");
  // monthly teaching hours (pending) - mthp
  reportSheet.getRange(mthpFirstRow + mthpRows - 1, mthpMonthColumn).setValue("Total");
  // monthly teaching hours (total) - mtht
  reportSheet.getRange(mthtFirstRow + mthtRows - 1, mthtMonthColumn).setValue("Total");
  // average weekly teaching hours (confirmed) - wthc
  reportSheet.getRange(wthcFirstRow + wthcRows - 1, wthcMonthColumn).setValue("Total");
  // average weekly teaching hours (pending) - wthp
  reportSheet.getRange(wthpFirstRow + wthpRows - 1, wthpMonthColumn).setValue("Total");
  // average weekly teaching hours (total) - wtht
  reportSheet.getRange(wthtFirstRow + wthtRows - 1, wthtMonthColumn).setValue("Total");
  
  // set borders
  // monthly teaching hours (confirmed) - mthc
  reportSheet.getRange(mthcFirstRow, mthcFirstColumn, mthcRows, mthcColumns).setBorder(true, true, true, true, true, true);
  reportSheet.getRange(mthcFirstRow, mthcFirstColumn, mthcRows, mthcColumns).setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.DOUBLE);
  // monthly teaching hours (pending) - mthp
  reportSheet.getRange(mthpFirstRow, mthpFirstColumn, mthpRows, mthpColumns).setBorder(true, true, true, true, true, true);
  reportSheet.getRange(mthpFirstRow, mthpFirstColumn, mthpRows, mthpColumns).setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.DOUBLE);
  // monthly teaching hours (total) - mtht
  reportSheet.getRange(mthtFirstRow, mthtFirstColumn, mthtRows, mthtColumns).setBorder(true, true, true, true, true, true);
  reportSheet.getRange(mthtFirstRow, mthtFirstColumn, mthtRows, mthtColumns).setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.DOUBLE);
  // average weekly teaching hours (confirmed) - wthc
  reportSheet.getRange(wthcFirstRow, wthcFirstColumn, wthcRows, wthcColumns).setBorder(true, true, true, true, true, true);
  reportSheet.getRange(wthcFirstRow, wthcFirstColumn, wthcRows, wthcColumns).setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.DOUBLE);
  // average weekly teaching hours (pending) - wthp
  reportSheet.getRange(wthpFirstRow, wthpFirstColumn, wthpRows, wthpColumns).setBorder(true, true, true, true, true, true);
  reportSheet.getRange(wthpFirstRow, wthpFirstColumn, wthpRows, wthpColumns).setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.DOUBLE);
  // average weekly teaching hours (total) - wtht
  reportSheet.getRange(wthtFirstRow, wthtFirstColumn, wthtRows, wthtColumns).setBorder(true, true, true, true, true, true);
  reportSheet.getRange(wthtFirstRow, wthtFirstColumn, wthtRows, wthtColumns).setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.DOUBLE);
  // attendence - att
  reportSheet.getRange(attFirstRow, attFirstColumn, attRows, attColumns).setBorder(true, true, true, true, true, true);
  reportSheet.getRange(attFirstRow, attFirstColumn, attRows, attColumns).setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.DOUBLE);
  
  // data calculation
  // get data
  var allData = classSheet.getRange(START_ROW_CS, START_COLUMN_CS, classSheet.getLastRow() - START_ROW_CS + 1, END_COLUMN_CS - START_COLUMN_CS + 1).getValues();
  // keep tracking during loop  
  var recordedConfirmedClassNum = 0;
  var recordedPendingClassNum = 0;
  var totalStudentNum = 0;
  // for each class
  for (var x = 0; x < allData.length; x++) {
    // skip this class or not?
    if (allData[x][START_DATE_COLUMN - 1] == "TBD" || allData[x][END_DATE_COLUMN - 1] == "TBD" || !allData[x][START_DATE_COLUMN - 1] || !allData[x][END_DATE_COLUMN - 1] || !allData[x][WEEK_HR_COLUMN - 1]) {
      continue;
    }
    // add student number to attendence
    if (allData[x][STUDENT_NUM_COLUMM - 1] && allData[x][STATUS_COLUMN - 1] == "Confirmed") {
      totalStudentNum += allData[x][STUDENT_NUM_COLUMM - 1];
      var targetCell = reportSheet.getRange(attHeadRow + 1 + reportClasses.indexOf(allData[x][LEVEL_COLUMN - 1]), attStudentNumColumn);
      targetCell.setValue(targetCell.getValue() + allData[x][STUDENT_NUM_COLUMM - 1]);
    }
    // monthly teaching hours (confirmed) - mthc
    if (allData[x][STATUS_COLUMN - 1] == "Confirmed") {
      var startDate = allData[x][START_DATE_COLUMN - 1];
      var endDate = allData[x][END_DATE_COLUMN - 1];
      //var serviceMonthsIndex = getMonthsIndex(startDate, endDate);
      var monthlyHours = new Array(reportMonths.length);
      for (var y = 0; y < monthlyHours.length; y++) {
        monthlyHours[y] = 0;
      }
      var weeklyDate = new Date(startDate.getTime());
      while (weeklyDate <= endDate) {
        var month = weeklyDate.getMonth(); 
        monthlyHours[reportMonthsIndex.indexOf(month)] += durationToSecond(allData[x][WEEK_HR_COLUMN - 1]);
        weeklyDate.setDate(weeklyDate.getDate() + 7);
      }
      // add monthly hours into reports
      var instructor = allData[x][INSTRUCTOR_COLUMN - 1];
      for (var y = 0; y < reportMonths.length; y++) {
        // monthly teaching hours (confirmed) - mthc
        var targetCell = reportSheet.getRange(mthcInstructorRow + 1 + y, mthcMonthColumn + 1 + reportInstructors.indexOf(getFirstName(instructor)));
        targetCell.setValue(targetCell.getValue() + monthlyHours[y]);
        // monthly teaching hours (total) - mtht
        targetCell = reportSheet.getRange(mthtInstructorRow + 1 + y, mthtMonthColumn + 1 + reportInstructors.indexOf(getFirstName(instructor)));
        targetCell.setValue(targetCell.getValue() + monthlyHours[y]);
      }
      var ta = allData[x][TA_COLUMN - 1];
      if (ta) {
        for (var y = 0; y < reportMonths.length; y++) {
          // monthly teaching hours (confirmed) - mthc
          var targetCell = reportSheet.getRange(mthcInstructorRow + 1 + y, mthcMonthColumn + 1 + reportInstructors.indexOf(getFirstName(ta)));
          targetCell.setValue(targetCell.getValue() + monthlyHours[y]);
          // monthly teaching hours (total) - mtht
          targetCell = reportSheet.getRange(mthtInstructorRow + 1 + y, mthtMonthColumn + 1 + reportInstructors.indexOf(getFirstName(ta)));
          targetCell.setValue(targetCell.getValue() + monthlyHours[y]);
        }
      }
      recordedConfirmedClassNum++;
    }
    // monthly teaching hours (pending) - mthp
    if (allData[x][STATUS_COLUMN - 1] == "Pending") {
      var startDate = allData[x][START_DATE_COLUMN - 1];
      var endDate = allData[x][END_DATE_COLUMN - 1];
      //var serviceMonthsIndex = getMonthsIndex(startDate, endDate);
      var monthlyHours = new Array(reportMonths.length);
      for (var y = 0; y < monthlyHours.length; y++) {
        monthlyHours[y] = 0;
      }
      var weeklyDate = new Date(startDate.getTime());
      while (weeklyDate <= endDate) {
        var month = weeklyDate.getMonth(); 
        monthlyHours[reportMonthsIndex.indexOf(month)] += durationToSecond(allData[x][WEEK_HR_COLUMN - 1]);
        weeklyDate.setDate(weeklyDate.getDate() + 7);
      }
      // add monthly hours into reports
      var instructor = allData[x][INSTRUCTOR_COLUMN - 1];
      for (var y = 0; y < reportMonths.length; y++) {
        // monthly teaching hours (pending) - mthp
        var targetCell = reportSheet.getRange(mthpInstructorRow + 1 + y, mthpMonthColumn + 1 + reportInstructors.indexOf(getFirstName(instructor)));
        targetCell.setValue(targetCell.getValue() + monthlyHours[y]);
        // monthly teaching hours (total) - mtht
        targetCell = reportSheet.getRange(mthtInstructorRow + 1 + y, mthtMonthColumn + 1 + reportInstructors.indexOf(getFirstName(instructor)));
        targetCell.setValue(targetCell.getValue() + monthlyHours[y]);
      }
      var ta = allData[x][TA_COLUMN - 1];
      if (ta) {
        for (var y = 0; y < reportMonths.length; y++) {
          // monthly teaching hours (pending) - mthp
          var targetCell = reportSheet.getRange(mthpInstructorRow + 1 + y, mthpMonthColumn + 1 + reportInstructors.indexOf(getFirstName(ta)));
          targetCell.setValue(targetCell.getValue() + monthlyHours[y]);
          // monthly teaching hours (total) - mtht
          targetCell = reportSheet.getRange(mthtInstructorRow + 1 + y, mthtMonthColumn + 1 + reportInstructors.indexOf(getFirstName(ta)));
          targetCell.setValue(targetCell.getValue() + monthlyHours[y]);
        }
      }
      recordedPendingClassNum++;
    }
  }
  
  // add total
  // monthly teaching hours (confirmed) - mthc
  if (recordedConfirmedClassNum) {
    var totalHoursPerMonth = new Array(reportMonths.length);
    var totalHoursPerInstructor = new Array(reportInstructors.length);
    for (var x = 0; x < totalHoursPerMonth.length; x++) {
      totalHoursPerMonth[x] = 0;
      var hoursPerInstructor = reportSheet.getRange(mthcInstructorRow + 1 + x, mthcMonthColumn + 1, 1, reportInstructors.length).getValues();
      for (var y = 0; y < hoursPerInstructor[0].length; y++) {
        totalHoursPerMonth[x] += hoursPerInstructor[0][y];
      }
      reportSheet.getRange(mthcInstructorRow + 1 + x, mthcMonthColumn + reportInstructors.length + 1).setValue(totalHoursPerMonth[x]);
    }
    for (var x = 0; x < totalHoursPerInstructor.length; x++) {
      totalHoursPerInstructor[x] = 0;
      var hoursPerMonth = reportSheet.getRange(mthcInstructorRow + 1, mthcMonthColumn + 1 + x, reportMonths.length, 1).getValues();
      for (var y = 0; y < hoursPerMonth.length; y++) {
        totalHoursPerInstructor[x] += hoursPerMonth[y][0];
      }
      reportSheet.getRange(mthcInstructorRow + reportMonths.length + 1, mthcMonthColumn + 1 + x).setValue(totalHoursPerInstructor[x]);
    }
  }
  // monthly teaching hours (pending) - mthp
  if (recordedPendingClassNum) {
    var totalHoursPerMonth = new Array(reportMonths.length);
    var totalHoursPerInstructor = new Array(reportInstructors.length);
    for (var x = 0; x < totalHoursPerMonth.length; x++) {
      totalHoursPerMonth[x] = 0;
      var hoursPerInstructor = reportSheet.getRange(mthpInstructorRow + 1 + x, mthpMonthColumn + 1, 1, reportInstructors.length).getValues();
      for (var y = 0; y < hoursPerInstructor[0].length; y++) {
        totalHoursPerMonth[x] += hoursPerInstructor[0][y];
      }
      reportSheet.getRange(mthpInstructorRow + 1 + x, mthpMonthColumn + reportInstructors.length + 1).setValue(totalHoursPerMonth[x]);
    }
    for (var x = 0; x < totalHoursPerInstructor.length; x++) {
      totalHoursPerInstructor[x] = 0;
      var hoursPerMonth = reportSheet.getRange(mthpInstructorRow + 1, mthpMonthColumn + 1 + x, reportMonths.length, 1).getValues();
      for (var y = 0; y < hoursPerMonth.length; y++) {
        totalHoursPerInstructor[x] += hoursPerMonth[y][0];
      }
      reportSheet.getRange(mthpInstructorRow + reportMonths.length + 1, mthpMonthColumn + 1 + x).setValue(totalHoursPerInstructor[x]);
    }
  }
  // monthly teaching hours (total) - mtht
  if (recordedConfirmedClassNum || recordedPendingClassNum) {
    var totalHoursPerMonth = new Array(reportMonths.length);
    var totalHoursPerInstructor = new Array(reportInstructors.length);
    for (var x = 0; x < totalHoursPerMonth.length; x++) {
      totalHoursPerMonth[x] = 0;
      var hoursPerInstructor = reportSheet.getRange(mthtInstructorRow + 1 + x, mthtMonthColumn + 1, 1, reportInstructors.length).getValues();
      for (var y = 0; y < hoursPerInstructor[0].length; y++) {
        totalHoursPerMonth[x] += hoursPerInstructor[0][y];
      }
      reportSheet.getRange(mthtInstructorRow + 1 + x, mthtMonthColumn + reportInstructors.length + 1).setValue(totalHoursPerMonth[x]);
    }
    for (var x = 0; x < totalHoursPerInstructor.length; x++) {
      totalHoursPerInstructor[x] = 0;
      var hoursPerMonth = reportSheet.getRange(mthtInstructorRow + 1, mthtMonthColumn + 1 + x, reportMonths.length, 1).getValues();
      for (var y = 0; y < hoursPerMonth.length; y++) {
        totalHoursPerInstructor[x] += hoursPerMonth[y][0];
      }
      reportSheet.getRange(mthtInstructorRow + reportMonths.length + 1, mthtMonthColumn + 1 + x).setValue(totalHoursPerInstructor[x]);
    }
  }
  
  // fill out average weekly hours reports
  // average weekly teaching hours (confirmed) - wthc
  if (recordedConfirmedClassNum) {
    var wthcDataRange = reportSheet.getRange(wthcInstructorRow + 1, wthcMonthColumn + 1, reportMonths.length + 1, reportInstructors.length + 1);
    var mthcDataRange = reportSheet.getRange(mthcInstructorRow + 1, mthcMonthColumn + 1, reportMonths.length + 1, reportInstructors.length + 1);
    var wthcvalues = wthcDataRange.getValues();
    var mthcvalues = mthcDataRange.getValues();
    for (var x = 0; x < wthcvalues.length; x++) {
      for (var y = 0; y < wthcvalues[0].length; y++) {
        if (mthcvalues[x][y]) {
          wthcvalues[x][y] = Math.round((mthcvalues[x][y] / 30) * 7);
          wthcvalues[x][y] = secondToDurationText(wthcvalues[x][y]);
        }
      }
    }
    wthcDataRange.setValues(wthcvalues);
  }
  // average weekly teaching hours (pending) - wthp
  if (recordedPendingClassNum) {
    var wthpDataRange = reportSheet.getRange(wthpInstructorRow + 1, wthpMonthColumn + 1, reportMonths.length + 1, reportInstructors.length + 1);
    var mthpDataRange = reportSheet.getRange(mthpInstructorRow + 1, mthpMonthColumn + 1, reportMonths.length + 1, reportInstructors.length + 1);
    var wthpvalues = wthpDataRange.getValues();
    var mthpvalues = mthpDataRange.getValues();
    for (var x = 0; x < wthpvalues.length; x++) {
      for (var y = 0; y < wthpvalues[0].length; y++) {
        if (mthpvalues[x][y]) {
          wthpvalues[x][y] = Math.round((mthpvalues[x][y] / 30) * 7);
          wthpvalues[x][y] = secondToDurationText(wthpvalues[x][y]);
        }
      }
    }
    wthpDataRange.setValues(wthpvalues);
  }
  // average weekly teaching hours (total) - wtht
  if (recordedConfirmedClassNum || recordedPendingClassNum) {
    var wthtDataRange = reportSheet.getRange(wthtInstructorRow + 1, wthtMonthColumn + 1, reportMonths.length + 1, reportInstructors.length + 1);
    var mthtDataRange = reportSheet.getRange(mthtInstructorRow + 1, mthtMonthColumn + 1, reportMonths.length + 1, reportInstructors.length + 1);
    var wthtvalues = wthtDataRange.getValues();
    var mthtvalues = mthtDataRange.getValues();
    for (var x = 0; x < wthtvalues.length; x++) {
      for (var y = 0; y < wthtvalues[0].length; y++) {
        if (mthtvalues[x][y]) {
          wthtvalues[x][y] = Math.round((mthtvalues[x][y] / 30) * 7);
          wthtvalues[x][y] = secondToDurationText(wthtvalues[x][y]);
        }
      }
    }
    wthtDataRange.setValues(wthtvalues);
  }
  // attendence - att
  reportSheet.getRange(attFirstRow + attRows - 1, attStudentNumColumn).setValue(totalStudentNum);
  
  // transform format
  // monthly teaching hours (confirmed) - mthc
  if (recordedConfirmedClassNum) {
    var hoursDataRange = reportSheet.getRange(mthcInstructorRow + 1, mthcMonthColumn + 1, reportMonths.length + 1, reportInstructors.length + 1);
    var values = hoursDataRange.getValues();
    for (var x = 0; x < values.length; x++) {
      for (var y = 0; y < values[0].length; y++) {
        if (values[x][y]) {
          values[x][y] = secondToDurationText(values[x][y]);
        }
      }
    }
    hoursDataRange.setValues(values);
  }
  // monthly teaching hours (pending) - mthp
  if (recordedPendingClassNum) {
    var hoursDataRange = reportSheet.getRange(mthpInstructorRow + 1, mthpMonthColumn + 1, reportMonths.length + 1, reportInstructors.length + 1);
    var values = hoursDataRange.getValues();
    for (var x = 0; x < values.length; x++) {
      for (var y = 0; y < values[0].length; y++) {
        if (values[x][y]) {
          values[x][y] = secondToDurationText(values[x][y]);
        }
      }
    }
    hoursDataRange.setValues(values);
  }
  // monthly teaching hours (total) - mtht
  if (recordedConfirmedClassNum || recordedPendingClassNum) {
    var hoursDataRange = reportSheet.getRange(mthtInstructorRow + 1, mthtMonthColumn + 1, reportMonths.length + 1, reportInstructors.length + 1);
    var values = hoursDataRange.getValues();
    for (var x = 0; x < values.length; x++) {
      for (var y = 0; y < values[0].length; y++) {
        if (values[x][y]) {
          values[x][y] = secondToDurationText(values[x][y]);
        }
      }
    }
    hoursDataRange.setValues(values);
  }
  // attendence - att
  if (reportClasses.length) {
    var range = reportSheet.getRange(attHeadRow + 1, attFirstColumn, reportClasses.length, 2);
    range.sort({column: attFirstColumn, ascending: true});
  }
}