function clearLogs() {
    var logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs")
    var lastLog = logSheet.getLastRow();
    var logRange = logSheet.getRange(2, 1, lastLog, 5);
    logRange.clear()
}

function writeLogs(student, email, location, program) {
    var logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs")
    logSheet.appendRow([new Date(), student, email, location, program]);
}

function alreadySent(student, email) {
    var logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs")
    var lastLog = logSheet.getLastRow() - 1;
    if (lastLog == 0) lastLog++
    var logRange = logSheet.getRange(2, 1, lastLog, 5);
    var logList = logRange.getValues();
    var res = false
    for (i in logList) {
        var studentLog = logList[i][1]
        var emailLog = logList[i][2]
        console.log({ student: email, sudentlog: studentLog })
        if (emailLog == email && studentLog == student) res = true
    }
    return res
}

function clearOldLogs() {
    var logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs")
    var lastLog = logSheet.getLastRow() - 1;
    var logRange = logSheet.getRange(2, 1, lastLog, 5);
    var logList = logRange.getValues();
    var res = false
    for (i in logList) {
        var date = logList[i][0]
        var dateObj = new Date(date)
        var then = dateObj.getTime()

        var now = new Date()
        var nowDate = now.getTime()
        var diff = nowDate - then
        var day = diff / 86400000
        var row = parseInt(i) + 2
        if (day > 30) {
            console.log('del row', row)
            logSheet.deleteRow(row)
        }
    }
}