function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("Ubiqum")

        .addItem("Refresh Totals", "getTotal")
        .addToUi();

}


function getTotal() {
    var relationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ubiqum")
    var startRow = 2;
    var startCol = 1
    var lastCourseLocation = relationSheet.getLastRow() - 1;
    var relationRange = relationSheet.getRange(startRow, startCol, lastCourseLocation, 3);
    var urlList = relationRange.getValues();

    for (i in urlList) {
        var ss = SpreadsheetApp.openByUrl(urlList[i][2]);
        var sheet = ss.getSheetByName("Students")

        var DATA = 0
        var MERN = 0
        var JAVA = 0

        var sRow = 2;
        var sCol = 1
        var lastStudent = sheet.getLastRow() - 1;
        var studentRange = sheet.getRange(sRow, sCol, lastStudent, 4);
        var studentList = studentRange.getValues();

        for (j in studentList) {
            var program = studentList[j][1]
            if (program == "DATA") DATA++
            if (program == "JAVA") JAVA++
            if (program == "MERN") MERN++

        }
        writeTotals(urlList[i][0], DATA, JAVA, MERN)
    }

}


function writeTotals(city, data, java, mern) {
    var totalsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Home")
    var lastlocation = totalsheet.getLastRow() - 1;
    var totalRange = totalsheet.getRange(2, 1, lastlocation, 4);
    var totalList = totalRange.getValues();

    for (i in totalList) {
        if (totalList[i][0] == city) {
            var row = parseInt(i) + 2
            totalsheet.getRange("B" + row).setValue(data)
            totalsheet.getRange("C" + row).setValue(java)
            totalsheet.getRange("D" + row).setValue(mern)

        }
    }

}

