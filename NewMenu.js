// The onOpen function is executed automatically every time a Spreadsheet is loaded
function onOpen() {
  var menuEntries = [];
  // When the user clicks on "addMenuExample" then "Menu Entry 1", the function function1 is
  // executed.
  menuEntries.push({name: "New Semester", functionName: "newSemester"});
  CLASS_SCHEDULE.addMenu("Operations", menuEntries);
}