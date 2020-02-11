function stripHTML(string) {
  return string.replace(/<[^>]+>/g, "").replace(/&nbsp;/g, " ").replace(/&ndash;/g, "-").replace(/\r/, "").replace(/\n/, "");
}

function setFormats() {
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dante's Workspace");
  mainSheet.getRange("A5:C1000").setNumberFormat("@");
  mainSheet.getRange("D5:E1000").setNumberFormat("mm/dd/yy");
  mainSheet.getRange("F5:U1000").setNumberFormat("@");
  mainSheet.getRange("G5:I1000").setHorizontalAlignment("right");
  mainSheet.hideColumns(15, 12);
}

function makeDate(string) {
  var isPM = string[string.length - 2] == "P";
  var times = string.replace(/AM/, "").replace(/PM/, "").replace(/ /, "").split(":");
  var d = new Date();
  if (isPM) {
    if (times[0] == "12") {
      d.setHours(12, parseInt(times[1]));
    } else {
      d.setHours(parseInt(times[0]) + 12, parseInt(times[1]));
    }
  } else {
    if (times[0] == "12") {
      d.setHours(0, parseInt(times[1]));
    } else {
      d.setHours(parseInt(times[0]), parseInt(times[1]));
    }
  }
  d.setFullYear(1970, 2, 1);
  d.setSeconds(0, 0);
  return d;
}

function main() {
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dante's Workspace");
  var curr_row = 5;
  var main_range = mainSheet.getRange("A5:Z1000");
  main_range.clearContent();
  main_range.clearNote();
  setFormats();

  var filter_day = mainSheet.getRange("C1").getValue();
  
  var search_location = mainSheet.getRange("B1").getValue();
  var converted_search_loc = "";
  if (search_location == "All Locations") {

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
  var chk_range = mainSheet.getRange("A5");
  if (chk_range.getValue() == "") {
    chk_range.setValue("Nothing found.");
  } else { // It is fine to proceed with sorting.
    var first_sort = mainSheet.getRange("B2").getValue();
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
      to_sort.push({column: 17, ascending: true});
      to_sort.push({column: 6, ascending: true});
    } else if (first_sort == "Day of week (SSFTWTM) and time (backwards)") {
      to_sort.push({column: 17, ascending: false});
      to_sort.push({column: 6, ascending: false});
    } else if (first_sort == "Most seats remaining") {
      to_sort.push({column: 18, ascending: false});
    } else if (first_sort == "Least seats remaining") {
      to_sort.push({column: 18, ascending: true});
    } else if (first_sort == "Largest capacity") {
      to_sort.push({column: 19, ascending: true});
    } else if (first_sort == "Smallest capacity") {
      to_sort.push({column: 19, ascending: false});
    } else if (first_sort == "Most students") {
      to_sort.push({column: 20, ascending: false});
    } else if (first_sort == "Least students") {
      to_sort.push({column: 20, ascending: true});
    } else if (first_sort == "Longest meeting duration") {
      to_sort.push({column: 21, ascending: false});
    } else if (first_sort == "Shortest meeting duration") {
      to_sort.push({column: 21, ascending: true});
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

    internal_set("A", value["address_name"]);
    internal_set("B", value["title"]);

    var class_day;
    if (value["monday"]) {
      class_day = "Monday";
      internal_set("Q", "0");
    } else if (value["tuesday"]) {
      class_day = "Tuesday";
      internal_set("Q", "1");
    } else if (value["wednesday"]) {
      class_day = "Wednesday";
      internal_set("Q", "2");
    } else if (value["thursday"]) {
      class_day = "Thursday";
      internal_set("Q", "3");
    } else if (value["friday"]) {
      class_day = "Friday";
      internal_set("Q", "4");
    } else if (value["saturday"]) {
      class_day = "Saturday";
      internal_set("Q", "5");
    } else if (value["sunday"]) {
      class_day = "Sunday";
      internal_set("Q", "6");
    } else {
      class_day = "?";
      internal_set("Q", "7");
    }
    internal_set("C", class_day); // No need to repeatedly override cell value
    internal_set("D", value["start_date"]);
    internal_set("E", value["end_date"]);
    internal_set("F", value["start_time"]);
    internal_set("G", value["end_time"]);
    var seatsTotal = value["class_size"];
    var seatsLeft = value["seats"];
    var seatsTaken = seatsTotal - seatsLeft;
    internal_set("H", seatsTaken.toString().concat("/".concat(seatsTotal.toString())));
    internal_set("I", value["ages"]);
    internal_set("J", "$".concat(parseInt(value["cost"]) + parseInt(value["charter_fee"])));
    mainSheet.getRange("K".concat(row_to_edit)).setNote(value["prerequisites"]);
    mainSheet.getRange("L".concat(row_to_edit)).setNote(stripHTML(value["description"]));
    mainSheet.getRange("M".concat(row_to_edit)).setNote(value["address"].concat("\n".concat(value["city"].concat(", CA ".concat(value["zipcode"])))));
    mainSheet.getRange("N".concat(row_to_edit)).setNote(stripHTML(value["schedule_notes"]));
    internal_set("O", value["name"]);
    internal_set("P", value["id"]);
    internal_set("R", seatsLeft);
    internal_set("S", seatsTotal);
    internal_set("T", seatsTaken);
    var diff = makeDate(value["end_time"]) - makeDate(value["start_time"]);
    diff /= (1000 * 60);  // convert diff from ms to mins
    internal_set("U", diff);
    curr_row++;
  }
}
