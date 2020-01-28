function newSemester() {
  var newSheetName = Browser.inputBox("Create a new sheet", "Please enter a name for your new sheet:", Browser.Buttons.OK_CANCEL);
  if (newSheetName != "cancel") {
    var newSheet = TEMPLATE_SHEET.copyTo(CLASS_SCHEDULE);
    newSheet.setName(newSheetName).showSheet().activate();
    newSheet.deleteRows(START_ROW_CS, newSheet.getLastRow() - START_ROW_CS + 1);
    var protection = newSheet.protect();
    // protection.removeEditors(protection.getEditors());
    protection.setWarningOnly(true);
    // Log
    var operation = "Create New Semester";
    var description = "New Semester: "+newSheetName+" has been created.";
    log(operation, description, "");
  }
}

function newClass() {
  var classID = Browser.inputBox("Class ID", "Please enter the class ID for the class:", Browser.Buttons.OK_CANCEL);
  if (!classID) {
    Browser.msgBox("Warning", "Please enter a class ID.", Browser.Buttons.OK);
  }
  else if (classID != "cancel") {
    var activeSheet = CLASS_SCHEDULE.getActiveSheet();
    var targetRow = activeSheet.getLastRow() + 1;
    // set new class
    var sourceRange = TEMPLATE_SHEET.getRange(START_ROW_CS, START_COLUMN_CS, 1, END_COLUMN_CS - START_COLUMN_CS + 1);
    var sourceFunctions = sourceRange.getFormulasR1C1();
    var targetRange = activeSheet.getRange(targetRow, START_COLUMN_CS, 1, END_COLUMN_CS - START_COLUMN_CS + 1);
    sourceRange.copyTo(targetRange);
    targetRange.clearContent();
    activeSheet.getRange(targetRow, CLASS_ID_COLUMN).setValue(classID);
    activeSheet.getRange(targetRow, WEEK_HR_COLUMN).setFormulaR1C1(sourceFunctions[0][WEEK_HR_COLUMN - 1]);
    activeSheet.getRange(targetRow, DAY_INDEX_COLUMN).setFormulaR1C1(sourceFunctions[0][DAY_INDEX_COLUMN - 1]);
    activeSheet.getRange(targetRow, STATUS_INDEX_COLUMN).setFormulaR1C1(sourceFunctions[0][STATUS_INDEX_COLUMN - 1]);
    // unprotect class range
    var protections = activeSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    var unprotectedRanges = [];
    unprotectedRanges.push(getClassDataRange(activeSheet));
    protections[0].setUnprotectedRanges(unprotectedRanges);
    /*
    activeSheet.getRange(targetRow, SELECT_COLUMN).insertCheckboxes();
    activeSheet.getRange(targetRow, CLASS_ID_COLUMN).setValue(classID);
    var dateValidationRule1 = SpreadsheetApp.newDataValidation().requireValueInList(DAYS, true).build();
    activeSheet.getRange(targetRow, DAY_COLUMN).setDataValidation(dateValidationRule1);
    var dateValidationRule2 = SpreadsheetApp.newDataValidation().requireValueInList(getClassLevel(), true).build();
    activeSheet.getRange(targetRow, LEVEL_COLUMN).setDataValidation(dateValidationRule2);
    var dateValidationRule3 = SpreadsheetApp.newDataValidation().requireValueInList(getInstructors(), true).build();
    activeSheet.getRange(targetRow, INSTRUCTOR_COLUMN).setDataValidation(dateValidationRule3);
    activeSheet.getRange(targetRow, TA_COLUMN).setDataValidation(dateValidationRule3);
    var dateValidationRule4 = SpreadsheetApp.newDataValidation().requireValueInList(getStatus(), true).build();
    activeSheet.getRange(targetRow, STATUS_COLUMN).setDataValidation(dateValidationRule4);
    var dateValidationRule5 = SpreadsheetApp.newDataValidation().requireValueInList(getAccountTypes(), true).build();
    activeSheet.getRange(targetRow, ACCOUNT_TYPE_COLUMN).setDataValidation(dateValidationRule5);
    activeSheet.getRange(targetRow, WEEK_HR_COLUMN).setFormulaR1C1("R[0]C[" + (END_TIME_COLUMN - WEEK_HR_COLUMN) + "]-R[0]C[" + (START_TIME_COLUMN - WEEK_HR_COLUMN) + "]");
    activeSheet.getRange(targetRow, START_COLUMN_CS, 1, END_COLUMN_CS - START_COLUMN_CS + 1).setBorder(true, true, true, true, true, true);
    */
    // Log
    var operation = "Create New Class";
    var description = "New class: "+classID+" has been created.";
    log(operation, description, "");
  }  
}

function deleteSelectedClasses() {
  // get selected rows
  var confirmation = Browser.msgBox("Confirmation", "Are you sure you want to delete all selected classes?", Browser.Buttons.YES_NO);
  if (confirmation == "no") { return; }
  var activeSheet = CLASS_SCHEDULE.getActiveSheet();
  var selectStatus = activeSheet.getRange(START_ROW_CS, SELECT_COLUMN, activeSheet.getLastRow() - START_ROW_CS + 1, 1).getValues();
  var selectedClassNum = 0;
  for (var x = 0; x < selectStatus.length; x++) {
    if (selectStatus[x][0]) {
      selectedClassNum++;
    }
  }
  if (selectedClassNum == 0) {
    Browser.msgBox("Warning", "You didn't select any classes!", Browser.Buttons.OK);
  }
  else {
    for (var x = 0; x < selectedClassNum; x++) {
      selectStatus = activeSheet.getRange(START_ROW_CS, SELECT_COLUMN, activeSheet.getLastRow() - START_ROW_CS + 1, 1).getValues();
      for (var y = 0; y < selectStatus.length; y++) {
        if (selectStatus[y][0]) {
          var targetRow = y + START_ROW_CS;
          var classId = activeSheet.getRange(targetRow, CLASS_ID_COLUMN).getValue();
          if (eventsExist(targetRow)) {
            deleteEvents(targetRow);
          }
          activeSheet.deleteRow(targetRow);
          // Log
          var operation = "Delete Class";
          var description = "Class: "+classId+" has been deleted from sheet.";
          log(operation, description, "");
          break;
        }
      }
    }
    // unprotect class range
    var protections = activeSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    var unprotectedRanges = [];
    unprotectedRanges.push(getClassDataRange(activeSheet));
    protections[0].setUnprotectedRanges(unprotectedRanges);
    Browser.msgBox("Success", "Selected classes have been deleted!", Browser.Buttons.OK);
  }
}



function sort() {
  var activeSheet = CLASS_SCHEDULE.getActiveSheet();
  var sortRange = activeSheet.getRange(START_ROW_CS, START_COLUMN_CS, activeSheet.getLastRow() - START_ROW_CS + 1, END_COLUMN_CS - START_COLUMN_CS + 1);
  sortRange.sort([{column: DAY_INDEX_COLUMN, ascending: true}, {column: STATUS_INDEX_COLUMN, ascending: true}, {column: START_TIME_COLUMN, ascending: true}]);
}

function sortByInstructor() {
  sort();
  var activeSheet = CLASS_SCHEDULE.getActiveSheet();
  var sortRange = getClassDataRange(activeSheet);
  sortRange.sort([{column: INSTRUCTOR_COLUMN, ascending: true}]);
}

// Select all classes
function selectAll() {
  var activeSheet = CLASS_SCHEDULE.getActiveSheet();
  var lastRow = activeSheet.getLastRow();
  var selectRange = activeSheet.getRange(START_ROW_CS, SELECT_COLUMN, lastRow - START_ROW_CS + 1);
  var selectValue = selectRange.getValues();
  for (var x = 0; x < selectValue.length; x++) {
    selectValue[x][0] = true;
  }
  selectRange.setValues(selectValue);
}

// Select no class
function selectNone() {
  var activeSheet = CLASS_SCHEDULE.getActiveSheet();
  var lastRow = activeSheet.getLastRow();
  var selectRange = activeSheet.getRange(START_ROW_CS, SELECT_COLUMN, lastRow - START_ROW_CS + 1);
  var selectValue = selectRange.getValues();
  for (var x = 0; x < selectValue.length; x++) {
    selectValue[x][0] = false;
  }
  selectRange.setValues(selectValue);
}

// Hide cancelled classes
function hideCancelled() {
  var activeSheet = CLASS_SCHEDULE.getActiveSheet();
  var lastRow = activeSheet.getLastRow();
  if (lastRow < START_ROW_CS) { return; }
  for (var i = START_ROW_CS; i <= lastRow; i++) {
    var status = activeSheet.getRange(i, STATUS_COLUMN).getValue();
    if (status == "Cancelled") {
      activeSheet.getRange(i, SELECT_COLUMN).setValue("false");
      activeSheet.hideRows(i);
    }
  }
}