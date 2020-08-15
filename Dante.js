function stripHTML(string) {
    return string.replace(/<[^>]+>/g, "").replace(/&nbsp;/g, " ").replace(/&ndash;/g, "-").replace(/\r\n/, "\n").replace("/&amp;/g", "&").trim();
}

function cleanAddress(addr) {
    return addr.replace(/\r\n/g, "\n").replace(/\n/g, ", ");
}

function setFormats() {
    var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dante's Workspace");
    mainSheet.getRange("A7:D1000").setNumberFormat("@");
    mainSheet.getRange("E7:F1000").setNumberFormat("mm/dd/yy");
    mainSheet.getRange("G7:Z1000").setNumberFormat("@");
    mainSheet.getRange("K7:L1000").setNumberFormat("$#,####0.00");
    mainSheet.getRange("E7:Z1000").setHorizontalAlignment("right");
    mainSheet.getRange("A7:D1000").setHorizontalAlignment("left");
    mainSheet.getRange("T7:Z1000").setHorizontalAlignment("left");
    mainSheet.hideColumns(20, 7);
}

function makeDate(string) {
    var isPM = string[string.length - 2] == "P";
    var times = string.replace(/AM/, "").replace(/PM/, "").replace(/ /, "").split(":");
    var d = new Date();
    if (isPM) {
        if (times[0] == "12") {
            d.setUTCHours(12, parseInt(times[1]));
        } else {
            d.setUTCHours(parseInt(times[0]) + 12, parseInt(times[1]));
        }
    } else {
        if (times[0] == "12") {
            d.setUTCHours(0, parseInt(times[1]));
        } else {
            d.setUTCHours(parseInt(times[0]), parseInt(times[1]));
        }
    }
    d.setFullYear(1970, 2, 1);
    d.setSeconds(0, 0);
    return d;
}

function fetchClass(id) {
    return JSON.parse(UrlFetchApp.fetch("https://enroll.barnabasrobotics.com/courses/".concat(id).concat("/info.json")).getContentText());
}

function makeXMXD(d) {
    return (d.getUTCMonth() + 1).toString().concat("/").concat(d.getUTCDate().toString());
}

function makeMMDD(date) {
    return Utilities.formatDate(date, "UTC", "MM/dd");
}

function fixTime(date, time_) {
    var date = date.trim();
    var time_ = time_.trim();
    return new Date(date + " " + time_ + " -0000");
}

function fullClassList() {
    /*
    Convert Barnabas and Bookwhen classes to an identical format.
    Returns a list of class dicts.
    Structure:
    id: string (both)
    source: string, (both)
    address: string plaintext address of course (both)
    description: string, Course info (both)
    name: string, letter and numbers, (barnabas only)
    title: string, course name (both)
    start_date: Date (both)
    end_date: Date (both)
    meeting_dates: list of mm/dd strings (both)
    no_class_dates: list of mm/dd strings of extrapolated dates where class does not meet when it should (both)
    teachers: list of teacher names (barnabas only)
    notes: string, additional information (barnabas only)
    ages: string, age range of class (barnabas only)
    tickets: list of ticket dicts (both)
        id: string, ticket ID (bookwhen only)
        name: string, ticket name (bookwhen only)
        seats: int, number of slots total
        details: string, expanded info about this ticket (bookwhen only)
        attendees: int, number of current enrollements
        cost: int cost of ticket
        instructional_cost: int cost of instruction (barnabas only)
        charter_fee: int cost for charter school students (barnabas only)
        kit_fee: int cost for kits (optional barnabas only)
        hsc_fee: int referral fee when offering classes through another buisness (optional barnabas only)
        availible_from: optional Date, start date of when the ticket is availible (bookwhen only)
        availible_to: optional Date, end date of when the ticket is availible (both)
    */

    function noClassDates(start_date, end_date, meeting_dates) {
        if (!(meeting_dates.length)) {
            return [];
        } else {
            var first_day = new Date(meeting_dates[0]).getUTCDay();
            for (var i = 0 ; i < meeting_dates.length; i++) {
                if (new Date(meeting_dates[i]).getUTCDay() !== first_day) {
                    return [];
                }
            }
        }

        var first_meeting = start_date;
        var current_meeting = new Date(first_meeting.getTime());
        var no_cls_dates = [];

        while (current_meeting.getTime() <= end_date.getTime()) {  // Not sure whether you can compare Date directly
            if (!meeting_dates.includes(makeXMXD(current_meeting))) {
                if (!no_cls_dates.includes(makeMMDD(current_meeting))) {
                    no_cls_dates.push(makeMMDD(current_meeting));
                }
            }
            current_meeting = new Date(current_meeting.getTime() + 7 * 86400000); // Classes that don't only meet every 7th day don't work yet
        }
        if (start_date.getUTCMonth() != end_date.getUTCMonth()) {
            return;
        }
        while (current_meeting.getUTCMonth() == start_date.getUTCMonth()) {
            if (!no_cls_dates.includes(makeMMDD(current_meeting))) {
                no_cls_dates.push(makeMMDD(current_meeting));
            }
            current_meeting = new Date(current_meeting.getTime() + 7 * 86400000);
        }
        current_meeting = new Date(current_meeting.getTime() + 7 * 86400000);
        current_meeting.setUTCDate(1);
        while (current_meeting.getUTCDay() != first_meeting.getUTCDay()) {
            current_meeting = new Date(current_meeting.getTime() + 86400000);
        }
        while (current_meeting.getTime() < first_meeting.getTime()) {
            if (!no_cls_dates.includes(makeMMDD(current_meeting))) {
                no_cls_dates.push(makeMMDD(current_meeting));
            }
            current_meeting = new Date(current_meeting.getTime() + 7 * 86400000);
        }

        return no_cls_dates;
    }

    var barnabas = JSON.parse(UrlFetchApp.fetch("https://enroll.barnabasrobotics.com/courses.json?search%5Bcity%5D=").getContentText());
    var options = {};
    options.headers = {"Authorization": "Basic " + Utilities.base64Encode("qw482dhb13ujduu46n02kyl2lpcc:")};
    var bookwhen = JSON.parse(UrlFetchApp.fetch("https://api.bookwhen.com/v2/events?include=tickets,tickets.events,location", options).getContentText());
    var super_events = bookwhen["data"];
    var included = bookwhen["included"];

    var final = [];

    function addBarnabasClass(cls, index) {
        var info = JSON.parse(UrlFetchApp.fetch(
            "https://enroll.barnabasrobotics.com/courses/".concat(cls["id"]).concat("/info.json")
        ).getContentText());
        var sched = JSON.parse(UrlFetchApp.fetch(
            "https://enroll.barnabasrobotics.com/courses/".concat(cls["id"]).concat("/schedule.json")
        ).getContentText());

        var ret = {};
        ret["source"] = "barnabas";
        ret["address"] = info["address"].trim() + " " + info["city"] + ", CA " + info["zipcode"];
        ret["id"] = info["id"].toString();
        ret["name"] = info["name"];
        ret["title"] = info["title"];
        ret["description"] = info["description"];
        ret["start_date"] = fixTime(info["start_date"], info["start_time"]);
        ret["end_date"] = fixTime(info["end_date"], info["end_time"]);
        ret["meeting_dates"] = sched["session_dates"];
        ret["no_class_dates"] = noClassDates(new Date(info["start_date"]), new Date(info["end_date"]), sched["session_dates"]);
        ret["teachers"] = info["teacher_list"];
        ret["notes"] = stripHTML(sched["notes"]);
        ret["ages"] = info["ages"].trim();  // sometimes comes with whitespace
        var pseudo_ticket = {};
        pseudo_ticket["seats"] = info["class_size"];
        pseudo_ticket["attendees"] = info["class_size"] - info["seats"];
        pseudo_ticket["cost"] = parseInt(info["cost"]);
        pseudo_ticket["instructional_cost"] = parseInt(info["instructional_fee"]);
        pseudo_ticket["charter_fee"] = parseInt(info["charter_fee"]);
        pseudo_ticket["availible_to"] = new Date(info["cancel_deadline"]);
        if (!(info["kit_fee"] == null)) {
            pseudo_ticket["kit_fee"] = info["kit_fee"];
        }
        if (!(info["hsc_fee"] == null)) {
            pseudo_ticket["hsc_fee"] = info["hsc_fee"];
        }
        ret["tickets"] = [pseudo_ticket];
        final.push(ret);
    }

    function addBookWhenClass(cls, index) {
        var attrs = cls["attributes"];
        var new_dict = {};

        var sub_events = [];
        var mini_id = cls["id"].split("-")[1];
        for (var i = 0; i < included.length; i++) {
            if (included[i]["type"] == "event") {
                var sub_mini_id = included[i]["id"].split("-")[1];
                if (mini_id == sub_mini_id) {
                    sub_events.push(included[i]);
                }
            }
        }

        new_dict["id"] = cls["id"];
        new_dict["source"] = "bookwhen";
        var loc_id = cls["relationships"]["location"]["data"]["id"];
        var loc;
        for (var i = 0; i < included.length; i++) {
            if (included[i]["id"] == loc_id) {
                loc = included[i];
            }
        }
        var addr = cleanAddress(loc["attributes"]["address_text"].trim());
        if (!(loc["attributes"]["additional_info"] == null || loc["attributes"]["additional_info"] == "")) {
            addr += "\n" + loc["attributes"]["additional_info"];
        }
        new_dict["address"] = addr;
        new_dict["description"] = attrs["details"];
        new_dict["title"] = attrs["title"];
        new_dict["start_date"] = new Date(attrs["start_at"].substring(0, attrs["start_at"].length - 6) + "-00:00");
        new_dict["end_date"] = new Date(attrs["end_at"].substring(0, attrs["end_at"].length - 6) + "-00:00");
        var session_dates = [];
        for (var i = 0; i < sub_events.length; i++) {
            var date = new Date(sub_events[i]["attributes"]["start_at"]);
            session_dates.push(makeMMDD(date));
        }
        new_dict["meeting_dates"] = session_dates;
        new_dict["no_class_dates"] = noClassDates(new Date(cls["start_at"]), new Date(cls["end_at"]), session_dates);

        var ticketobjs = cls["relationships"]["tickets"]["data"];
        var ticket_ids = [];
        for (var i = 0; i < ticketobjs.length; i++) {
            ticket_ids.push(ticketobjs[i]["id"]);
        }
        ticketobjs = [];
        for (var i = 0; i < included.length; i++) {
            if (ticket_ids.includes(included[i]["id"])) {
                ticketobjs.push(included[i]);
            }
        }
        var tickets_to_push = [];
        for (var i = 0; i < ticketobjs.length; i++) {
            var running_ticket = ticketobjs[i];
            var new_ticket = {};
            new_ticket["id"] = running_ticket["id"];
            new_ticket["name"] = running_ticket["attributes"]["title"];
            new_ticket["seats"] = running_ticket["attributes"]["group_max"];
            new_ticket["details"] = running_ticket["attributes"]["details"];
            new_ticket["attendees"] = running_ticket["attributes"]["number_taken"];
            new_ticket["cost"] = running_ticket["attributes"]["cost"]["net"] / 100;
            var availible_from = running_ticket["attributes"]["availible_from"];
            var availible_to = running_ticket["attributes"]["availible_to"];
            if (!(availible_from == null)) {
                new_ticket["availible_from"] = new Date(availible_from);
            }
            if (!(availible_to == null)) {
                new_ticket["availible_to"] = new Date(availible_to);
            }
            tickets_to_push.push(new_ticket);
        }
        new_dict["tickets"] = tickets_to_push;
        final.push(new_dict);
    }

    for (var i = 0; i < barnabas.length; i++) {
        barnabas[i][1].forEach(addBarnabasClass);
    }
    super_events.forEach(addBookWhenClass);

    return final;
}


function main() {
    var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dante's Workspace");
    var main_range = mainSheet.getRange("A7:Z1000");
    var protection = mainSheet.protect().setDescription("Loading...");
    var chk_range = mainSheet.getRange("A7");
    protection.setWarningOnly(true);
    var curr_row = 7;
    main_range.clearContent();
    main_range.clearNote();
    chk_range.setValue("Constructing additional pylons...");
    setFormats();

    var filter_day = mainSheet.getRange("B3").getValue();

    var filter_dur = mainSheet.getRange("B1").getValue();
    var converted_filter_dur = 0;
    if (filter_dur == "1-hour/55-minute classes") {
        converted_filter_dur = 1;
    } else if (filter_dur == "90-minute classes") {
        converted_filter_dur = 2;
    } else if (filter_dur == "2-hour classes") {
        converted_filter_dur = 3;
    } else if (filter_dur == ">2-hour classes") {
        converted_filter_dur = 4;
    }

    var filter_enrollment = mainSheet.getRange("B2").getValue();
    var converted_filter_enrollment = 0;
    if (filter_enrollment == "Empty classes") {
        converted_filter_enrollment = 1;
    } else if (filter_enrollment == "Partially filled classes") {
        converted_filter_enrollment = 2;
    } else if (filter_enrollment == "Full classes") {
        converted_filter_enrollment = 3;
    }

    var resp = fullClassList();

    resp.forEach(addToSheet);

    // At this point, if A7:Z1000 is empty or default, no results were returned.
    if (chk_range.getValue() == "Constructing additional pylons..." || chk_range.getValue() == "") {
      chk_range.setValue("Nothing found.");
    } else { // It is fine to proceed with sorting.
        var first_sort = mainSheet.getRange("B4").getValue();
        var to_sort = [];
        if (first_sort == "Location") {
            to_sort.push({column: 1, ascending: true});
        } else if (first_sort == "Title (A to Z)") {
            to_sort.push({column: 2, ascending: true});
        } else if (first_sort == "Title (Z to A)") {
            to_sort.push({column: 2, ascending: false});
        } else if (first_sort == "Day of week (MTWTFSS) and time (forwards)") {
            to_sort.push({column: 20, ascending: true});
            to_sort.push({column: 7, ascending: true});
        } else if (first_sort == "Day of week (SSFTWTM) and time (backwards)") {
            to_sort.push({column: 20, ascending: false});
            to_sort.push({column: 7, ascending: false});
        } else if (first_sort == "Highest cost") {
            to_sort.push({column: 11, ascending: false});
        } else if (first_sort == "Lowest cost") {
            to_sort.push({column: 11, ascending: true});
        } else if (first_sort == "Highest expected revenue") {
            to_sort.push({column: 12, ascending: false});
        } else if (first_sort == "Lowest expected revenue") {
            to_sort.push({column: 12, ascending: true});
        } else if (first_sort == "Earliest cancellation deadline") {
            to_sort.push({column: 13, ascending: true});
        } else if (first_sort == "Latest cancellation deadline") {
            to_sort.push({column: 13, ascending: false});
        } else if (first_sort == "Most seats remaining") {
            to_sort.push({column: 21, ascending: false});
        } else if (first_sort == "Least seats remaining") {
            to_sort.push({column: 21, ascending: true});
        } else if (first_sort == "Largest capacity") {
            to_sort.push({column: 22, ascending: true});
        } else if (first_sort == "Smallest capacity") {
            to_sort.push({column: 22, ascending: false});
        } else if (first_sort == "Most students") {
            to_sort.push({column: 23, ascending: false});
        } else if (first_sort == "Least students") {
            to_sort.push({column: 23, ascending: true});
        } else if (first_sort == "Longest meeting duration") {
            to_sort.push({column: 24, ascending: false});
        } else if (first_sort == "Shortest meeting duration") {
            to_sort.push({column: 24, ascending: true});
        } else if (first_sort == "Source (Barnabas site first)") {
            to_sort.push({column: 25, ascending: true});
        } else if (first_sort == "Source (Bookwhen first)") {
            to_sort.push({column: 25, ascending: false});
        }

        var second_sort = mainSheet.getRange("B7").getValue();
        if (second_sort == "Location") {
            to_sort.push({column: 1, ascending: true});
        } else if (second_sort == "Title (A to Z)") {
            to_sort.push({column: 2, ascending: true});
        } else if (second_sort == "Title (Z to A)") {
            to_sort.push({column: 2, ascending: false});
        } else if (second_sort == "Day of week (MTWTFSS) and time (forwards)") {
            to_sort.push({column: 20, ascending: true});
            to_sort.push({column: 7, ascending: true});
        } else if (second_sort == "Day of week (SSFTWTM) and time (backwards)") {
            to_sort.push({column: 20, ascending: false});
            to_sort.push({column: 7, ascending: false});
        } else if (second_sort == "Highest cost") {
            to_sort.push({column: 11, ascending: false});
        } else if (second_sort == "Lowest cost") {
            to_sort.push({column: 11, ascending: true});
        } else if (second_sort == "Highest expected revenue") {
            to_sort.push({column: 12, ascending: false});
        } else if (second_sort == "Lowest expected revenue") {
            to_sort.push({column: 12, ascending: true});
        } else if (second_sort == "Earliest cancellation deadline") {
            to_sort.push({column: 13, ascending: true});
        } else if (second_sort == "Latest cancellation deadline") {
            to_sort.push({column: 13, ascending: false});
        } else if (second_sort == "Most seats remaining") {
            to_sort.push({column: 21, ascending: false});
        } else if (second_sort == "Least seats remaining") {
            to_sort.push({column: 21, ascending: true});
        } else if (second_sort == "Largest capacity") {
            to_sort.push({column: 22, ascending: true});
        } else if (second_sort == "Smallest capacity") {
            to_sort.push({column: 22, ascending: false});
        } else if (second_sort == "Most students") {
            to_sort.push({column: 23, ascending: false});
        } else if (second_sort == "Least students") {
            to_sort.push({column: 23, ascending: true});
        } else if (second_sort == "Longest meeting duration") {
            to_sort.push({column: 24, ascending: false});
        } else if (second_sort == "Shortest meeting duration") {
            to_sort.push({column: 24, ascending: true});
        } else if (second_sort == "Source (Barnabas site first)") {
            to_sort.push({column: 25, ascending: true});
        } else if (second_sort == "Source (Bookwhen first)") {
            to_sort.push({column: 25, ascending: false});
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

    function addToSheet(classInfo, _index) {
        var row_to_edit = curr_row.toString();

        // This will be used later to set the column used to sort by duration, but we set it here for filtering
        // purposes.
        var end_time = new Date(classInfo["end_date"]);
        end_time.setUTCDate(classInfo["start_date"].getUTCDate());
        end_time.setUTCMonth(classInfo["start_date"].getUTCMonth());
        var class_duration = (classInfo["end_date"] - classInfo["start_date"]) % (1000 * 60 * 60 * 24);
        class_duration /= (1000 * 60);  // convert ms to mins

        switch (converted_filter_dur) {
            case 1:
                if (class_duration < 55 || class_duration > 60) {
                    return;
                } else {
                    break;
                }
            case 2:
                if (class_duration < 85 || class_duration > 95) {
                    return;
                } else {
                    break;
                }
            case 3:
                if (class_duration < 105 || class_duration > 125) {
                    return;
                } else {
                    break;
                }
            case 4:
                if (class_duration < 125) {
                    return;
                } else {
                    break;
                }
            default:
                break;
        }

        var days_of_week = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "???"];
        var day_id = new Date(classInfo["meeting_dates"][0]).getUTCDay();
        var all_days = [];
        var class_day = days_of_week[day_id];
        if (day_id == NaN) {
            var class_day = "???";
            day_id = -1;
        } else {
            var class_day = days_of_week[day_id];
        }

        for (var i = 0; i < classInfo["meeting_dates"].length; i++) {
            var day = new Date(classInfo["meeting_dates"][i]).getUTCDay();
            if (day != NaN && !(all_days.includes(days_of_week[day]))) {
                all_days.push(days_of_week[day]);
            }
        }

        function dateStuff() {
            mainSheet.getRange("C".concat(row_to_edit)).setValue(class_day);
            mainSheet.getRange("T".concat(row_to_edit)).setValue(day_id.toString());
            if (all_days.length > 1) {
                internal_set("O", "??");
                return;
            }
            if (classInfo["no_class_dates"] == null || classInfo["no_class_dates"].length == 0) {
                internal_set("O", "??");
                return;
            } else {
                mainSheet.getRange("O".concat(row_to_edit)).setValue(classInfo["no_class_dates"].join(", "));
            }
        }

        // These will be used later for sorting and class size column. We set them here for filtering.
        var seatsTotal = 0;
        var seatsTaken = 0;
        function add_ticket_to_total(t, i) {
            seatsTotal += t["seats"];
            seatsTaken += t["attendees"];
        }
        classInfo["tickets"].forEach(add_ticket_to_total);
        if (seatsTaken > seatsTotal) {
            var swap = seatsTaken;
            seatsTaken = seatsTotal;
            seatsTotal = swap;
        }
        var seatsLeft = seatsTotal - seatsTaken;


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
            if (!(all_days.includes(filter_day))) {
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

        internal_set("A", classInfo["address"]);
        internal_set("B", classInfo["title"]);
        dateStuff(classInfo, row_to_edit);
        if (!(classInfo["teachers"] == null)) {
            internal_set("D", classInfo["teachers"].join(", "));
        } else {
            internal_set("D", "None found.");
        }
        internal_set("E", makeXMXD(classInfo["start_date"]));
        internal_set("F", makeXMXD(classInfo["end_date"]));
        internal_set("G", Utilities.formatDate(new Date(classInfo["start_date"]), "UTC", "hh:mm a"));
        internal_set("H", Utilities.formatDate(new Date(classInfo["end_date"]), "UTC", "hh:mm a"));
        internal_set("I", seatsTaken.toString().concat("/".concat(seatsTotal.toString())));
        if (!(classInfo["ages"] == null)) {
            internal_set("J", classInfo["ages"]);
        } else {
            internal_set("J", "?");
        }
        var first_ticket = classInfo["tickets"][0];
        var charter_fee = first_ticket["charter_fee"]
        if (charter_fee == null) {
            charter_fee = 0;
        }
        internal_set("K", "$".concat(first_ticket["cost"] + charter_fee));
        internal_set("L", "$".concat(seatsTaken * first_ticket["cost"] + charter_fee));
        if (!(first_ticket["availible_to"] == null)) {
            internal_set("M", makeXMXD(first_ticket["availible_to"]));
        }
        mainSheet.getRange("N".concat(row_to_edit)).setNote(stripHTML(classInfo["description"]));
        internal_set("N", "Hover:");
        // O set by dateStuff
        if (!(classInfo["notes"] == null)) {
            internal_set("P", stripHTML(classInfo["notes"]));
        }
        internal_set("Q", classInfo["meeting_dates"].join(", "));
        internal_set("R", classInfo["name"]);
        internal_set("S", classInfo["id"]);
        // Call to dateStuff changes row T
        internal_set("U", seatsLeft);
        internal_set("V", seatsTotal);
        internal_set("W", seatsTaken);
        internal_set("X", class_duration);
        internal_set("Z", classInfo["source"]);
        curr_row++;
    }
}

function generateReport() {
    var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dante's Hours Test");

    var month = "3";  // March
    var resp = JSON.parse(UrlFetchApp.fetch("https://enroll.barnabasrobotics.com/courses.json?search%5Bcity%5D=").getContentText());

    var classes = [];
    for (var i = 0; i < resp.length; i++) {
        for (var j = 0; j < resp[i][1].length; j++) {
            classes.push(resp[i][1][j]);
        }
    }

    var new_classes = [];
    for (var i = 0; i < classes.length; i++) {
        schedule = JSON.parse(UrlFetchApp.fetch("https://enroll.barnabasrobotics.com/courses/".concat(classes[i]["id"]).concat("/schedule.json")).getContentText());
        var found = false;
        for (var j = 0; j < schedule["session_dates"].length; j++){
            if (schedule["session_dates"][j].split("/")[0] == month) {
                found = true;
            }
        }
        if (found) {
            classes[i]["schedule"] = schedule;
            classes[i]["info"] = JSON.parse(UrlFetchApp.fetch("https://enroll.barnabasrobotics.com/courses/".concat(classes[i]["id"]).concat("/info.json")).getContentText());
            new_classes.push(classes[i]);
        }
    }

    var hours = {};

    for (var i = 0; i < new_classes.length; i++) {
        var teachers = new_classes[i]["info"]["teacher_list"];

        var meetings = new_classes[i]["schedule"]["session_dates"];
        var actual_meetings = [];
        for (var j = 0; j < meetings.length; j++) {
            var current_meeting = meetings[j];
            if (current_meeting.split("/")[0] == month) {
                actual_meetings.push(current_meeting);
            }
        }

        var meeting_duration = makeDate(new_classes[i]["end_time"]) - makeDate(new_classes[i]["start_time"]);
        meeting_duration /= 1000 * 60 * 60;
        for (var j = 0; j < teachers.length; j++) {
            if (hours[teachers[j]] == undefined) {
                hours[teachers[j]] = [0, 0];
            }
            var to_add = meeting_duration * actual_meetings.length;
            to_add = [Math.floor(to_add), 60 * (to_add % 1)];

            hours[teachers[j]][0] += to_add[0];
            hours[teachers[j]][1] += to_add[1];
        }
    }

    mainSheet.getRange("A2").setValue(hours);

    // var to_write = []
    
    // var keys = hours.keys();

    // for (var i = 0; i < keys.length; i++) {
    //     to_write.push([keys[i], hours[keys[i]]])
    // }

    // mainSheet.getRange("A2:B100").setValues(to_write);
}
