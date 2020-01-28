function stripHTML(string) {
  return string.replace(/<[^>]+>/g, "").replace(/&nbsp;/g, " ").replace(/&ndash;/g, "-");
}

function setFormats() {
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dante's Workspace");
  mainSheet.getRange("A4:C1000").setNumberFormat("@");
  mainSheet.getRange("D4:E1000").setNumberFormat("mm/dd/yy");
  mainSheet.getRange("F4:P1000").setNumberFormat("@");
  mainSheet.getRange("G3:H1000").setHorizontalAlignment("right");
  mainSheet.getRange("I3:I1000").setHorizontalAlignment("right");
  mainSheet.hideColumns(14, 13);
}

function main() {
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dante's Workspace");
  var curr_row = 4;
  var main_range = mainSheet.getRange("A4:Z1000");
  main_range.clearContent();
  main_range.clearNote();
  setFormats();

  var filter_day = mainSheet.getRange("A1").getValue();
  
  var search_location = mainSheet.getRange("B1").getValue();
  var converted_search_loc = "";
  if (search_location == "All Locations") {

  } else if (search_location == "Anaheim (Cavalry Baptist)") {
    converted_search_loc = "Anaheim";
  } else if (search_location == "Arcadia (FUNdamentals)") {
    converted_search_loc = "Arcadia";
  } else if (search_location == "Monrovia (EIE)") {
    converted_search_loc = "Monrovia";
  } else if (search_location == "Pasadena (Barnabas HQ)") {
    converted_search_loc = "Pasadena";
  } else {
    // There is some sort of invalid input here.
    // Assume all locations are to be searched for.
  }

  var resp = JSON.parse(UrlFetchApp.fetch("https://enroll.barnabasrobotics.com/courses.json?search%5Bcity%5D=".concat(converted_search_loc)).getContentText());

  for (var i = 0; i < resp.length; i++) {
      resp[i][1].forEach(addToSheet);
  }

  // At this point, if A3:Z1000 is empty, no results were returned.
  var chk_range = mainSheet.getRange("A4")
  if (chk_range.getValue() == "") {
    chk_range.setValue("Nothing found.")
  } else { // It is fine to proceed with sorting.
    var first_sort = mainSheet.getRange("A2").getValue();
    var to_sort = [];
    if (first_sort == "Highest cost") {
      to_sort.push({column: 10, ascending: false});
    } else if (first_sort == "Lowest cost") {
      to_sort.push({column: 10, ascending: true});
    } else if (first_sort == "Title (A to Z)") {
      to_sort.push({column: 2, ascending: true});
    } else if (first_sort == "Title (Z to A)") {
      to_sort.push({column: 2, ascending: false});
    } else if (first_sort == "Location") {
      to_sort.push({column: 3, ascending: true});
    } else if (first_sort == "Day of week (MTWTFSS) and time (forwards)") {
      to_sort.push({column: 16, ascending: true});
      to_sort.push({column: 6, ascending: true});
    } else if (first_sort == "Day of week (SSFTWTM) and time (backwards)") {
      to_sort.push({column: 16, ascending: false});
      to_sort.push({column: 6, ascending: false});
    }
    main_range.sort(to_sort);
  }

  function addToSheet(value, index) {
    var row_to_edit = curr_row.toString();

    function internal_set(col, val) {
      var to_append = ""
      if (val == null  || val == undefined) {
        to_append = "?";
      } else {
        to_append = val.toString();
      }
      mainSheet.getRange(col.concat(row_to_edit)).setValue(to_append);

    }

    if (filter_day != "All Days") {
      if (!(value[filter_day.toLowerCase()])) {
        return;
      }
    }

    var class_day;
    if (value["monday"]) {
      class_day = "Monday";
      internal_set("P", "0");
    } else if (value["tuesday"]) {
      class_day = "Tuesday";
      internal_set("P", "1");
    } else if (value["wednesday"]) {
      class_day = "Wednesday";
      internal_set("P", "2");
    } else if (value["thursday"]) {
      class_day = "Thursday";
      internal_set("P", "3");
    } else if (value["friday"]) {
      class_day = "Friday";
      internal_set("P", "4");
    } else if (value["saturday"]) {
      class_day = "Saturday";
      internal_set("P", "5");
    } else if (value["sunday"]) {
      class_day = "Sunday";
      internal_set("P", "6");
    } else {
      class_day = "?";
      internal_set("P", "7");
    }
    internal_set("A", class_day); // No need to repeatedly override cell value
    internal_set("B", value["title"]);
    internal_set("C", value["address_name"]);
    internal_set("D", value["start_date"]);
    internal_set("E", value["end_date"]);
    internal_set("F", value["start_time"]);
    internal_set("G", value["end_time"]);
    internal_set("H", (value["class_size"] - value["seats"]).toString().concat("/".concat(value["class_size"].toString())));
    internal_set("I", value["ages"]);
    internal_set("J", "$".concat(parseInt(value["cost"]) + parseInt(value["charter_fee"])));
    mainSheet.getRange("K".concat(row_to_edit)).setNote(value["prerequisites"]);
    mainSheet.getRange("L".concat(row_to_edit)).setNote(stripHTML(value["description"]));
    mainSheet.getRange("M".concat(row_to_edit)).setNote(value["address"].concat("\n".concat(value["city"].concat(", CA ".concat(value["zipcode"])))));
    internal_set("N", value["name"]);
    internal_set("O", value["id"]);

    curr_row++;
  }
}
