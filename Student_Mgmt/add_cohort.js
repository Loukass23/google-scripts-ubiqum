function addCohort() {
    var cohortSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cohorts")
    var studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students")
    var lastStudent = studentsSheet.getLastRow() - 1;

    //get program/cohort/start date/students number values
    var program = cohortSheet.getRange("M2").getValue()
    var cohortN = cohortSheet.getRange("N2").getValue()
    var startDate = cohortSheet.getRange("O2").getValue()
    var studentsN = cohortSheet.getRange("P2").getValue()

    //define which program the new cohort should follow
    var programSheet = getProgramSheet(program)

    //get cohort cells
    var programPlanRange = programSheet.getRange(2, 2, 1, 3);
    var programPlanList = programPlanRange.getValues();
    var module = programPlanList[0][0]
    var sprint = programPlanList[0][1]
    var day = programPlanList[0][2]
    var programSpikeRange = programSheet.getRange("E2");
    var programSpikeValue = programSpikeRange.getValue()
    //add new raw for cohort

    var lastCohort = cohortSheet.getLastRow() + 1;

    var formula =
        '=IF(B' + lastCohort + '="MERN"  ,WORKDAY.INTL(D' + lastCohort + ',59,1,Holidays!$A:$A),WORKDAY.INTL(D' + lastCohort + ',99,1,Holidays!$A:$A))'

    var refDayFormula = '=NETWORKDAYS.INTL(D' + lastCohort + ',TODAY(),1,Holidays!$A:$A)'

    cohortSheet.appendRow(["", program, cohortN, startDate, studentsN, module, sprint, day, 1, formula, programSpikeValue, refDayFormula]);




    //add new raws for students
    for (var i = 0; i < studentsN; i++) {

        var lastStudent = studentsSheet.getLastRow() + 1;
        var formulaJustified = '=COUNTIF(Q' + lastStudent + ':FH' + lastStudent + ',"Justified")/COUNTA(Q' + lastStudent + ':' + lastStudent + ')'
        var formulaYES = '=COUNTIF(Q' + lastStudent + ':FH' + lastStudent + ', "Yes")/COUNTA(Q' + lastStudent + ':' + lastStudent + ')'
        var formulaNo = '=COUNTIF(Q' + lastStudent + ':FH' + lastStudent + ', "No")/COUNTA(Q' + lastStudent + ':' + lastStudent + ')'
        var refDayFormula = '=NETWORKDAYS.INTL(P' + lastStudent + ',TODAY(),1,Holidays!$A:$A)'

        studentsSheet.appendRow(["", program, cohortN, "Firstname", "initial.name.ubiqum@gmail.com", module, sprint, day, 1, refDayFormula, "add link", "add link", formulaJustified, formulaYES, formulaNo, startDate]);
    }
}
