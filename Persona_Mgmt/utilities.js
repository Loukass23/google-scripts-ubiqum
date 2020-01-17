function onOpen() {

    var ui = SpreadsheetApp.getUi();
    ui.createMenu("Ubiqum")
        .addItem("Set triggers", "setTriggers")
        .addItem("Cancel triggers", "cancelTriggers")
        .addItem("Manual Send", "mainController")
        .addItem("Clear Old Logs", "clearOldLogs")
        .addItem("Clear Logs", "clearLogs")
        .addToUi();

};

function setTriggers() {
    cancelTriggers()
    var emailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Emails")

    var startRow = 2;
    var startCol = 7
    var lastEmail = emailSheet.getLastRow() - 1;
    var emailRange = emailSheet.getRange(startRow, startCol, lastEmail, 1);
    var emailList = emailRange.getValues();
    var duplicates = []
    for (j in emailList) {
        var email = emailList[j]

        if (duplicates.indexOf(email[0]) == -1) {
            duplicates.push(email[0])

            ScriptApp.newTrigger("mainController")
                .timeBased()
                .atHour(email[0])
                .everyDays(1) // Frequency is required if you are using atHour() or nearMinute()
                .create();
        }
    }
}

function cancelTriggers() {

    var triggers = ScriptApp.getProjectTriggers();

    for (var i = 0; i < triggers.length; i++) {

        ScriptApp.deleteTrigger(triggers[i]);

    };
};
