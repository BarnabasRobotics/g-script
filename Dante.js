function stripHTML(string) {
    return string.replace(/<[^>]+>/g, "").replace(/&nbsp;/g, " ").replace(/&ndash;/g, "-").replace(/\r/, "").replace(/\n/, "");
}

function setFormats() {
    var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dante's Workspace");
    mainSheet.getRange("A9:D1000").setNumberFormat("@");
    mainSheet.getRange("E9:F1000").setNumberFormat("mm/dd/yy");
    mainSheet.getRange("G9:Z1000").setNumberFormat("@");
    mainSheet.getRange("M9:M1000").setNumberFormat("mm/dd/yy");
    mainSheet.getRange("E9:Z1000").setHorizontalAlignment("right");
    mainSheet.getRange("A9:D1000").setHorizontalAlignment("left");
    mainSheet.getRange("Q9:R1000").setHorizontalAlignment("left");
    mainSheet.hideColumns(19, 8);
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

function fetchClass(id) {
    return JSON.parse(UrlFetchApp.fetch("https://enroll.barnabasrobotics.com/courses/".concat(id).concat("/info.json")).getContentText());
}

function convertDate(d) {
    return (d.getUTCMonth() + 1).toString().concat("/").concat(d.getUTCDate());
}

function main() {
    var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dante's Workspace");
    var main_range = mainSheet.getRange("A9:Z1000");
    var protection = mainSheet.protect().setDescription("Loading...");
    protection.setWarningOnly(true);
    var curr_row = 9;
    main_range.clearContent();
    main_range.clearNote();
    setFormats();

    var filter_day = mainSheet.getRange("B4").getValue();

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

    var filter_teacher = mainSheet.getRange("B5").getValue();
    var converted_filter_teacher = "*";
    if (filter_teacher == "Ed") {
        converted_filter_teacher = "Edward Li";
    } else if (filter_teacher == "Eric") {
        converted_filter_teacher = "Eric Lin";
    } else if (filter_teacher == "Petra") {
        converted_filter_teacher = "Petra Poschmann";
    }

    var filter_dur = mainSheet.getRange("B2").getValue();
    var converted_filter_dur = 0;
    if (filter_dur == "1-hour/55-minute classes") {
        converted_filter_dur = [60, 55];
    } else if (filter_dur == "90-minute classes") {
        converted_filter_dur = [90];
    } else if (filter_dur == "2-hour classes") {
        converted_filter_dur = [120];
    }

    var filter_enrollment = mainSheet.getRange("B3").getValue();
    var converted_filter_enrollment = 0;
    if (filter_enrollment == "Empty classes") {
        converted_filter_enrollment = 1;
    } else if (filter_enrollment == "Partially filled classes") {
        converted_filter_enrollment = 2;
    } else if (filter_enrollment == "Full classes") {
        converted_filter_enrollment = 3;
    }

    var resp = JSON.parse(UrlFetchApp.fetch("https://enroll.barnabasrobotics.com/courses.json?search%5Bcity%5D=".concat(converted_search_loc)).getContentText());

    for (var i = 0; i < resp.length; i++) {
        resp[i][1].forEach(addToSheet);
    }

    // At this point, if A3:Z1000 is empty, no results were returned.
    var chk_range = mainSheet.getRange("A9");
    if (chk_range.getValue() == "") {
      chk_range.setValue("Nothing found.");
    } else { // It is fine to proceed with sorting.
        var first_sort = mainSheet.getRange("B6").getValue();
        var to_sort = [];
        if (first_sort == "Title (A to Z)") {
            to_sort.push({column: 2, ascending: true});
        } else if (first_sort == "Title (Z to A)") {
            to_sort.push({column: 2, ascending: false});
        } else if (first_sort == "Location") {
            to_sort.push({column: 3, ascending: true});
        } else if (first_sort == "Teacher") {
            to_sort.push({column: 4, ascending: true});
        } else if (first_sort == "Highest cost") {
            to_sort.push({column: 11, ascending: false});
        } else if (first_sort == "Lowest cost") {
            to_sort.push({column: 11, ascending: true});
        } else if (first_sort == "Highest expected revenue") {
            to_sort.push({column: 12, ascending: false});
        } else if (first_sort == "Lowest expected revenue") {
            to_sort.push({column: 12, ascending: true});
        } else if (first_sort == "Day of week (MTWTFSS) and time (forwards)") {
            to_sort.push({column: 21, ascending: true});
            to_sort.push({column: 7, ascending: true});
        } else if (first_sort == "Day of week (SSFTWTM) and time (backwards)") {
            to_sort.push({column: 21, ascending: false});
            to_sort.push({column: 7, ascending: false});
        } else if (first_sort == "Most seats remaining") {
            to_sort.push({column: 22, ascending: false});
        } else if (first_sort == "Least seats remaining") {
            to_sort.push({column: 22, ascending: true});
        } else if (first_sort == "Largest capacity") {
            to_sort.push({column: 23, ascending: true});
        } else if (first_sort == "Smallest capacity") {
            to_sort.push({column: 23, ascending: false});
        } else if (first_sort == "Most students") {
            to_sort.push({column: 24, ascending: false});
        } else if (first_sort == "Least students") {
            to_sort.push({column: 24, ascending: true});
        } else if (first_sort == "Longest meeting duration") {
            to_sort.push({column: 25, ascending: false});
        } else if (first_sort == "Shortest meeting duration") {
            to_sort.push({column: 25, ascending: true});
        } else if (first_sort == "Highest level") {
            to_sort.push({column: 26, ascending: false});
        } else if (first_sort == "Lowest level") {
            to_sort.push({column: 26, ascending: true});
        }

        var second_sort = mainSheet.getRange("B7").getValue();
        if (second_sort == "Title (A to Z)") {
            to_sort.push({column: 2, ascending: true});
        } else if (second_sort == "Title (Z to A)") {
            to_sort.push({column: 2, ascending: false});
        } else if (second_sort == "Location") {
            to_sort.push({column: 3, ascending: true});
        } else if (second_sort == "Teacher") {
            to_sort.push({column: 4, ascending: true});
        } else if (second_sort == "Highest cost") {
            to_sort.push({column: 11, ascending: false});
        } else if (second_sort == "Lowest cost") {
            to_sort.push({column: 11, ascending: true});
        } else if (second_sort == "Highest expected revenue") {
            to_sort.push({column: 12, ascending: false});
        } else if (second_sort == "Lowest expected revenue") {
            to_sort.push({column: 12, ascending: true});
        } else if (second_sort == "Day of week (MTWTFSS) and time (forwards)") {
            to_sort.push({column: 21, ascending: true});
            to_sort.push({column: 7, ascending: true});
        } else if (second_sort == "Day of week (SSFTWTM) and time (backwards)") {
            to_sort.push({column: 21, ascending: false});
            to_sort.push({column: 7, ascending: false});
        } else if (second_sort == "Most seats remaining") {
            to_sort.push({column: 22, ascending: false});
        } else if (second_sort == "Least seats remaining") {
            to_sort.push({column: 22, ascending: true});
        } else if (second_sort == "Largest capacity") {
            to_sort.push({column: 23, ascending: true});
        } else if (second_sort == "Smallest capacity") {
            to_sort.push({column: 23, ascending: false});
        } else if (second_sort == "Most students") {
            to_sort.push({column: 24, ascending: false});
        } else if (second_sort == "Least students") {
            to_sort.push({column: 24, ascending: true});
        } else if (second_sort == "Longest meeting duration") {
            to_sort.push({column: 25, ascending: false});
        } else if (second_sort == "Shortest meeting duration") {
            to_sort.push({column: 25, ascending: true});
        } else if (second_sort == "Highest level") {
            to_sort.push({column: 26, ascending: false});
        } else if (second_sort == "Lowest level") {
            to_sort.push({column: 26, ascending: true});
        }
        if (to_sort.length > 0) {
            main_range.sort(to_sort);
        }
    }

    protection.remove();
    other_protections = mainSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    for (var i = 0; i < other_protections.length; i++) {
        if (other_protections[i].canEdit() && other_protections[i].getDescription() == "Loading...") {
            other_protections[i].remove();
        }
    }

    function dateStuff(class_dict, row) {
        var class_day;
        if (class_dict["sunday"]) {
            class_day = "Sunday";
            day_id = 0;
        } else if (class_dict["monday"]) {
            class_day = "Monday";
            day_id = 1;
        } else if (class_dict["tuesday"]) {
            class_day = "Tuesday";
            day_id = 2;
        } else if (class_dict["wednesday"]) {
            class_day = "Wednesday";
            day_id = 3;
        } else if (class_dict["thursday"]) {
            class_day = "Thursday";
            day_id = 4;
        } else if (class_dict["friday"]) {
             class_day = "Friday";
             day_id = 5;
        } else if (class_dict["saturday"]) {
            class_day = "Saturday";
            day_id = 6;
        } else {
            class_day = "?";
            day_id = 7;
        }
        mainSheet.getRange("C".concat(row)).setValue(class_day);
        mainSheet.getRange("U".concat(row)).setValue(day_id.toString());

        var start_date = new Date(class_dict["start_date"]);
        var end_date = new Date(class_dict["end_date"]);

        meeting_dates = JSON.parse(UrlFetchApp.fetch(
            "https://enroll.barnabasrobotics.com/courses/".concat(class_dict["id"]).concat("/schedule.json")
        ).getContentText())["session_dates"];

        var first_meeting = new Date(start_date);
        // while (first_meeting.getDay() != day_id) {
        //     // first_meeting.setDay(first_meeting.getDate() + 1); leaves potential issues with the last day of
        //     // the month, as we might be trying to make a new date for September 32nd, for example.
        //     // instead we add 86400000 milliseconds, which is a day.
        //     first_meeting = new Date(first_meeting.getTime() + 86400000);
        // }

        var current_meeting = first_meeting;
        var no_cls_dates = [];

        while (current_meeting.getTime() <= end_date.getTime()) {  // Not sure whether you can compare Date directly
            if (meeting_dates.indexOf(convertDate(current_meeting)) == -1) {
                if (no_cls_dates.indexOf(convertDate(current_meeting)) == -1) {
                    no_cls_dates.push(convertDate(current_meeting));
                }
            }
            current_meeting = new Date(current_meeting.getTime() + 7 * 86400000); // Classes that don't only meet every 7th day don't work yet
        }
        function addExtraNoClassDates() {
            if (start_date.getUTCMonth() != end_date.getUTCMonth()) {
                return;
            }
            while (current_meeting.getUTCMonth() == start_date.getUTCMonth()) {
                no_cls_dates.push(convertDate(current_meeting));
                current_meeting = new Date(current_meeting.getTime() + 7 * 86400000);
            }
            current_meeting = new Date(current_meeting.getTime() - 7 * 86400000);
            current_meeting.setUTCDate(1);
            while (current_meeting.getUTCDay() != day_id) {
                current_meeting = new Date(current_meeting.getTime() + 86400000);
            }
            while (current_meeting.getTime() < first_meeting.getTime()) {
                no_cls_dates.push(convertDate(current_meeting));
                current_meeting = new Date(current_meeting.getTime() + 7 * 86400000);
            }
        }

        addExtraNoClassDates();
        mainSheet.getRange("Q".concat(row)).setValue(no_cls_dates.join(", "));
    }

    function addToSheet(value, _index) {
        var row_to_edit = curr_row.toString();
        classInfo = fetchClass(value["id"]);

        if (converted_filter_teacher != "*") {
            teachers = classInfo["teacher_list"];
            if (!teachers.includes(converted_filter_teacher)) {
                return;
            }
        }

        // This will be used later to set the column used to sort by duration, but we set it here for filtering
        // purposes.
        var class_duration = makeDate(classInfo["end_time"]) - makeDate(classInfo["start_time"]);
        class_duration /= (1000 * 60);  // convert ms to mins

        if (converted_filter_dur != 0) {
            if (!converted_filter_dur.includes(class_duration)) {
                return;
            }
        }

        // These will be used later for sorting and class size column. We set them here for filtering.
        var seatsTotal = classInfo["class_size"];
        var seatsLeft = classInfo["seats"];
        var seatsTaken = seatsTotal - seatsLeft;

        if (converted_filter_enrollment != 0) {
            if (seatsTaken != 0 && converted_filter_enrollment == 1) {
                return;
            }
            if (seatsLeft != 0 && converted_filter_enrollment == 3) {
                return;
            }
            if ((seatsLeft == 0 || seatsTaken == 0) && converted_filter_enrollment == 2) {
                return;
            }
        }

        if (filter_day != "All Days") {
            if (!(classInfo[filter_day.toLowerCase()])) {
                return;
            }
        }

        function internal_set(col, val) {
            var to_append = ""
            if (val == null  || val == undefined) {
                to_append = "?";
            } else {
                to_append = val.toString();
            }
            mainSheet.getRange(col.concat(row_to_edit)).setValue(to_append);
        }

        internal_set("A", classInfo["address_name"]);
        internal_set("B", classInfo["title"]);
        dateStuff(classInfo, row_to_edit);
        internal_set("D", classInfo["teacher_list"].join(", "));
        internal_set("E", classInfo["start_date"]);
        internal_set("F", classInfo["end_date"]);
        internal_set("G", classInfo["start_time"]);
        internal_set("H", classInfo["end_time"]);
        internal_set("I", seatsTaken.toString().concat("/".concat(seatsTotal.toString())));
        internal_set("J", classInfo["ages"]);
        internal_set("K", "$".concat(parseInt(classInfo["cost"]) + parseInt(classInfo["charter_fee"])));
        internal_set("L", "$".concat(parseInt(seatsTaken) * parseInt(classInfo["cost"])));
        internal_set("M", classInfo["cancel_deadline"]);
        mainSheet.getRange("N".concat(row_to_edit)).setNote(classInfo["prerequisites"]);
        mainSheet.getRange("O".concat(row_to_edit)).setNote(stripHTML(classInfo["description"]));
        mainSheet.getRange("P".concat(row_to_edit)).setNote(classInfo["address"].concat("\n".concat(classInfo["city"].concat(", CA ".concat(classInfo["zipcode"])))));
        // Q set by dateStuff
        internal_set("R", stripHTML(classInfo["schedule_notes"]));
        internal_set("S", classInfo["name"]);
        internal_set("T", classInfo["id"]);
        // Call to dateStuff changes row U
        internal_set("V", seatsLeft);
        internal_set("W", seatsTotal);
        internal_set("X", seatsTaken);
        internal_set("Y", class_duration);
        internal_set("Z", classInfo["level_id"]);
        curr_row++;
    }
}
