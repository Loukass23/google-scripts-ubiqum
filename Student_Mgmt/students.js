//set colum to current date and hide past
//function onOpen() {
//  var studentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students") 
//  var lastStudent = studentSheet.getLastRow() -1;
//var dateRange = studentSheet.getRange(1, 16, 1, studentSheet.getLastColumn())
//var values = dateRange.getValues();
//  values = values[0];
//  var day = 24*3600*1000;  
//  var today = parseInt((new Date().setHours(0,0,0,0))/day);  
//
//  var ssdate; 
//  for (var i=0; i<values.length; i++) {
//    try {
//     ssdate = values[i].getTime()/day;
//    }
//    catch(e) {
//    }
//    if (ssdate && Math.floor(ssdate) == today) {
//
//      var pastRange = dateRange.offset(0,i,1,1)
//      studentSheet.setActiveRange(dateRange.offset(0,i,lastStudent,1));
//      studentSheet.hideColumn(dateRange.offset(0,0,lastStudent,i))
//  
//    }    
// }
//  var ui = SpreadsheetApp.getUi();
//  ui.createMenu("Ubiqum")
//  .addItem("Increment Manually Day Sprints","triggerEvent") 
//
//
//  .addToUi();
//}


function getTodayColumn() {
    var studentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students")
    var lastStudent = studentSheet.getLastRow() - 1;
    var dateRange = studentSheet.getRange(1, 16, 1, studentSheet.getLastColumn())
    var values = dateRange.getValues();
    values = values[0];
    var day = 24 * 3600 * 1000;
    var today = parseInt((new Date().setHours(0, 0, 0, 0)) / day);

    var ssdate;
    for (var i = 0; i < values.length; i++) {
        try {
            ssdate = values[i].getTime() / day;
        }
        catch (e) {
        }
        if (ssdate && Math.floor(ssdate) == today) {

            var pastRange = dateRange.offset(0, i, 1, 1)

            return dateRange.offset(1, i, lastStudent, 1)

        }
    }
}

function setAttendancetoYes() {
    var studentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students")
    var lastStudent = studentSheet.getLastRow() - 1;
    //    var studentsRange = studentSheet.getRange(i, 2);

    var day = 24 * 3600 * 1000;
    var today = parseInt((new Date().setHours(0, 0, 0, 0)) / day);
    var attendanceRange = getTodayColumn()
    if (attendanceRange) {
        var attendanceValues = attendanceRange.getValues()
        var att = []
        for (var i = 0; i < attendanceValues.length; i++) {
            if (attendanceValues[i][0] == "") {
                att.push(["Yes"])
            }
            else att.push(attendanceValues[i])
        }
        attendanceRange.setValues(att)
    }
}
//function getColumnNrByName(sheet, name) {
//  var range = sheet.getRange(1, 1, 1, sheet.getMaxColumns());
//  var values = range.getValues();
//    var day = 24*3600*1000;  
//
//  for (var row in values) {
//    for (var col in values[row]) {
//      if (values[row][col].getTime()/day == name) {
//        return parseInt(col);
//      }
//    }
//  }
//  
//  throw 'failed to get column by name';
//}

function getStudent() {
    var studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");

    //get students range
    var startRowStudents = 2;
    var startColStudents = 6;
    var lastStudent = studentsSheet.getLastRow();

    for (var i = startRowStudents; i <= lastStudent; i++) {
        //getting students program JAVA/MERN cell
        var studentsProgramRange = studentsSheet.getRange(i, 2);
        var studentsProgramValue = studentsProgramRange.getValue();
        //getting students module/sprint/day cells
        var studentsRange = studentsSheet.getRange(i, startColStudents, 1, 3);
        var studentsList = studentsRange.getValues();
        //getting students day
        var studentsDayRange = studentsSheet.getRange("I" + i);
        var studentsDayValue = studentsDayRange.getValue();
        //getting students reference day cell
        var studentsRefDayRange = studentsSheet.getRange("J" + i);
        //var studentsRefDayValue = studentsRefDayRange.getValue();
        //increasing students day and students reference day by 1
        studentsDayRange.setValue(studentsDayValue + 1)
        //studentsRefDayRange.setValue(studentsRefDayValue + 1)
        var currentRowProgram = studentsDayValue + 1;
        // change all cells value by incrementing reference program row of 1


        incrementSprint(studentsRange, currentRowProgram, getProgramSheet(studentsProgramValue));

    }
}
function addDeltaDayToStudent() {
    var studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students")
    //get active cell and increment value by 1
    var currentCell = studentsSheet.getActiveCell()

    //get row of incremented cell

    var currentStudentRow = currentCell.getRow();
    var sprintDayCell = studentsSheet.getRange("I" + currentStudentRow);
    var currentRowProgram = sprintDayCell.getValue();

    sprintDayCell.setValue(sprintDayCell.getValue() + 1)
    var newCurrentRowProgram = sprintDayCell.getValue();
    //getting student program JAVA/MERN
    var studentProgramRange = studentsSheet.getRange(currentStudentRow, 2);
    var studentsProgramValue = studentProgramRange.getValue();
    //getting student module/sprint/day
    var startColStudent = 6;
    var studentRange = studentsSheet.getRange(currentStudentRow, startColStudent, 1, 3);
    var studentList = studentRange.getValues();
    //change all cells value by incrementing reference program row of 1  

    incrementSprint(studentRange, newCurrentRowProgram, getProgramSheet(studentsProgramValue));


}
function subtractDeltaDayToStudent() {
    var studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students")

    //get active cell and increment value by 1
    var currentCell = studentsSheet.getActiveCell()
    //  currentCell.setValue(currentCell.getValue() - 1)
    //get row of incremented cell

    var currentStudentRow = currentCell.getRow();
    var sprintDayCell = studentsSheet.getRange("I" + currentStudentRow);
    var currentRowProgram = sprintDayCell.getValue();

    sprintDayCell.setValue(sprintDayCell.getValue() - 1)
    var newCurrentRowProgram = sprintDayCell.getValue();

    //getting student program JAVA/MERN
    var studentProgramRange = studentsSheet.getRange("B" + currentStudentRow);
    var studentsProgramValue = studentProgramRange.getValue();
    //getting student module/sprint/day
    var startColStudent = 6;
    var studentRange = studentsSheet.getRange(currentStudentRow, startColStudent, 1, 3);
    var studentList = studentRange.getValues();
    //change all cells value by incrementing reference program row of 1  

    incrementSprint(studentRange, newCurrentRowProgram, getProgramSheet(studentsProgramValue));

}





function incrementSprint(studentsRange, currentRowProgram, programSheet) {
    var startColProgram = 2;
    var currentRowWithHeader = currentRowProgram + 1
    var programPlanRange = programSheet.getRange(currentRowWithHeader, startColProgram, 1, 3);
    var programPlanList = programPlanRange.getValues();
    var programSpikeRange = programSheet.getRange(currentRowWithHeader, 5);
    var programSpikeValue = programSpikeRange.getValue()
    studentsRange.setValues(programPlanList);
}
