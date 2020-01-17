function getProgramSheet(program) {
    var url = "https://docs.google.com/spreadsheets/d/1uYc8ytyybHLdQlWGUnrluh6ZusI7f_Ooum6g2SDy6eg/edit?usp=sharing"
    switch (program) {
        case "JAVA": return SpreadsheetApp.openByUrl(url).getSheetByName("JAVA Program");
            break;
        case "MERN": return SpreadsheetApp.openByUrl(url).getSheetByName("MERN Program");
            break;
        case "DATA": return SpreadsheetApp.openByUrl(url).getSheetByName("DATA Program");
            break;
        default: return console.log("error on program")
    }
}

function isHoliday(date) {
    //get today date

    var formattedTodayDate = Utilities.formatDate(date, "GMT+2", "MM-dd-yyyy");
    //get holiday dates
    var holidaysSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Holidays");
    var lastHoliday = holidaysSheet.getLastRow();
    var holidaysRange = holidaysSheet.getRange(1, 1, lastHoliday, 1)
    var holidaysValue = holidaysRange.getValues();
    var formattedHolidayDateArray = [];
    for (var i = 0; i < holidaysValue.length; i++) {
        var formattedHolidayDate = Utilities.formatDate(holidaysValue[i][0], "GMT+2", "MM-dd-yyyy")
        if (formattedHolidayDate == formattedTodayDate) return true
    }
    return false
}

function isWeekDay(date) {
    if (date.getDay() >= 1 && date.getDay() <= 5) return true
    else return false

}
function validWorkingDay(date) {
    if (isWeekDay(date) && !isHoliday(date)) return true
    else return false
}

//this function gets triggered every day between 6am and 7am
function triggerEvent() {
    var today = new Date();
    if (validWorkingDay(today)) {
        //increment spike cohorts and students sheet
        getSpike();
        getStudent();
    }
}

function getSpike() {
    var cohortSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cohorts");
    //get cohort range
    var startRowCohort = 2;
    var startColCohort = 6;
    var lastCohort = cohortSheet.getLastRow();
    for (var i = startRowCohort; i <= lastCohort; i++) {
        //getting cohort program JAVA/MERN cell
        var cohortsProgramRange = cohortSheet.getRange(i, 2);
        var cohortsProgramValue = cohortsProgramRange.getValue();

        //getting cohort module/sprint/day cells
        var cohortsRange = cohortSheet.getRange(i, startColCohort, 1, 3);
        var cohortsList = cohortsRange.getValues();
        //getting cohort spike cell
        var cohortSpikeRange = cohortSheet.getRange(i, 11, 1, 1);
        //getting cohort day
        var cohortsDayRange = cohortSheet.getRange(i, 9);
        var cohortsDayValue = cohortsDayRange.getValue();
        //getting cohort reference day cell
        var cohortsRefDayRange = cohortSheet.getRange("L" + i);
        //var cohortsRefDayValue = cohortsRefDayRange.getValue();
        //increasing cohort day and cohort reference day by 1
        cohortsDayRange.setValue(cohortsDayValue + 1)
        //cohortsRefDayRange.setValue(cohortsRefDayValue + 1)
        var currentRowProgram = cohortsDayValue + 1;
        // change all cells value by incrementing reference program row of 1 
        incrementValue(cohortsRange, cohortSpikeRange, currentRowProgram, getProgramSheet(cohortsProgramValue))
    }
}

function addDeltaDay() {
    var cohortSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cohorts")
    //get active cell and increment value by 1
    var currentCell = cohortSheet.getActiveCell()

    //get row of incremented cell


    var currentCohortRow = currentCell.getRow();
    var sprintDayCell = cohortSheet.getRange("I" + currentCohortRow);
    sprintDayCell.setValue(sprintDayCell.getValue() + 1)
    var currentRowProgram = sprintDayCell.getValue();
    //getting cohort program JAVA/MERN
    var cohortsProgramRange = cohortSheet.getRange("B" + currentCohortRow);
    var cohortsProgramValue = cohortsProgramRange.getValue();
    //getting cohort module/sprint/day
    var startColCohort = 6;
    var cohortsRange = cohortSheet.getRange(currentCohortRow, startColCohort, 1, 3);
    var cohortsList = cohortsRange.getValues();
    //getting cohort spike
    var cohortSpikeRange = cohortSheet.getRange(currentCohortRow, 11, 1, 1);
    //change all cells value by incrementing reference program row of 1  

    incrementValue(cohortsRange, cohortSpikeRange, currentRowProgram, getProgramSheet(cohortsProgramValue));

}
function subtractDeltaDay() {
    var cohortSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cohorts")

    var currentCell = cohortSheet.getActiveCell()


    var currentCohortRow = currentCell.getRow();
    var sprintDayCell = cohortSheet.getRange("I" + currentCohortRow);

    sprintDayCell.setValue(sprintDayCell.getValue() - 1)
    var currentRowProgram = sprintDayCell.getValue();
    //getting cohort program JAVA/MERN
    var cohortsProgramRange = cohortSheet.getRange("B" + currentCohortRow);
    var cohortsProgramValue = cohortsProgramRange.getValue();

    //getting cohort module/sprint/day
    var startColCohort = 6;
    var cohortsRange = cohortSheet.getRange(currentCohortRow, startColCohort, 1, 3);
    var cohortsList = cohortsRange.getValues();
    //getting cohort spike
    var cohortSpikeRange = cohortSheet.getRange(currentCohortRow, 11, 1, 1);
    //change all cells value by decreasing reference program row of 1  
    incrementValue(cohortsRange, cohortSpikeRange, currentRowProgram, getProgramSheet(cohortsProgramValue));
    //  getProgramSheet(cohortsProgramValue)

}

function incrementValue(cohortsRange, cohortSpikeRange, currentRowProgram, programSheet) {
    var startColProgram = 2;
    var currentRowWithHeader = currentRowProgram + 1
    var programPlanRange = programSheet.getRange(currentRowWithHeader, startColProgram, 1, 3);
    var programPlanList = programPlanRange.getValues();
    var programSpikeRange = programSheet.getRange(currentRowWithHeader, 5);
    var programSpikeValue = programSpikeRange.getValue()
    var formula = programSpikeRange.getFormula();

    cohortsRange.setValues(programPlanList);

    if (formula != "") cohortSpikeRange.setValue(formula)
    else cohortSpikeRange.setValue(programSpikeValue)
}
