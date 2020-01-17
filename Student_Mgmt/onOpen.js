function onOpen() {
    var studentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students")
    var lastStudent = studentSheet.getLastRow() - 1;
    var dateRange = studentSheet.getRange(1, 17, 1, studentSheet.getLastColumn())
    var values = dateRange.getValues();
    values = values[0];
    var day = 24 * 3600 * 1000;
    var today = parseInt((new Date().setHours(0, 0, 0, 0)) / day);
    Logger.log(today);
    var ssdate;
    for (var i = 0; i < values.length; i++) {
        try {
            ssdate = values[i].getTime() / day;
        }
        catch (e) {
        }
        if (ssdate && Math.floor(ssdate) == today) {

            var pastRange = dateRange.offset(0, i, 1, 1)
            studentSheet.setActiveRange(dateRange.offset(0, i, lastStudent, 1));
            studentSheet.hideColumn(dateRange.offset(0, 0, lastStudent, i))
            break;
        }
    }
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("Increment")
        .addItem("Increment Manually Day Sprints", "triggerEvent")
        .addItem("Set All Daily Attendance to Yes  ", "setAttendancetoYes")

        .addToUi();
    var uiStudent = SpreadsheetApp.getUi();
    ui.createMenu("Students")
        .addItem("Fast student", "addDeltaDayToStudent")
        .addItem("Slow student", "subtractDeltaDayToStudent")

        .addToUi();
    ui.createMenu("Cohorts")
        .addItem("Fast cohort", "addDeltaDay")
        .addItem("Slow cohort", "subtractDeltaDay")

        .addToUi();
}