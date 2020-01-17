function eventController() {
    var calendarSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar")
    var startRow = 3;
    var startCol = 1;
    var calendarRange = calendarSheet.getRange(startRow, startCol, 1, 11)
    var event = calendarRange.getValues()[0]

    var eventProgram = event[0]
    var eventCohort1 = event[1]
    var eventCohort2 = event[2]
    var name = event[3]
    var description = event[4]
    var date = event[5]
    var timeCell = event[6]
    var time = timeCell.split("h")
    var timeHour = parseInt(time[0])
    var timeMinute = parseInt(time[1])
    var duration = event[7].split("h")
    var durationHour = parseInt(duration[0])
    var durationMinute = parseInt(duration[1])
    var location = event[8]
    var type = event[9]
    var endDate = event[10]

    var startDate = new Date(date)
    startDate.setHours(timeHour, timeMinute)
    var endDuration = new Date(startDate)
    endDuration.setHours(startDate.getHours() + durationHour);
    endDuration.setMinutes(startDate.getMinutes() + durationMinute);

    var studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students")
    var startRow = 2;
    var startCol = 2
    var lastStudent = studentsSheet.getLastRow() - 1;
    var studentRange = studentsSheet.getRange(startRow, startCol, lastStudent, 4)
    var studentList = studentRange.getValues();

    var staffMail = getStaff(eventProgram)
    var studentMail = [];


    for (var i = 0; i < studentList.length; i++) {

        var student = studentList[i]
        if (student[3] !== "" || student[3] !== "initial.name.ubiqum@gmail.com") {
            if (eventProgram == "ALL") studentMail.push(student[3])
            else if (eventProgram == student[0] && eventCohort1 == student[1]) studentMail.push(student[3])
            else if (eventProgram == "FULL-STACK") {
                if ("JAVA" == student[0] && eventCohort1 == student[1] || "MERN" == student[0] && eventCohort2 == student[1]) {

                    studentMail.push(student[3])
                }
            }
        }
    }
    var fullList = studentMail.concat(staffMail)
    var fullStrList = fullList.join(",")
    var studentStrList = studentMail.join(",")


    if (type == "Spike") {

        addSpikeEvent(fullStrList, name, description, startDate, endDuration, location)
    }

    if (type == "Stand Up") {
        if (eventProgram == "FULL-STACK") {
            var name = eventProgram + "-JAVA" + eventCohort1 + "_MERN-" + eventCohort2;
        }
        else {
            var name = eventProgram + "-" + eventCohort1
        }

        addStandUpEvent(fullStrList, name, description, startDate, endDuration, location, endDate)
    }
    if (type == "Milestones") {

        addMilestoneEvents(fullList, eventProgram, eventCohort1, date)
    }
}

function getMilestoneDate(start_date, num_days, program) {
    var tempCalcCell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar").getRange("M3")
    var formatDate = Utilities.formatDate(new Date(start_date), "GMT+1", "MM/dd/yyyy")
    if (program == 'MERN') {
        var formula = '=WORKDAY.INTL("' + formatDate + '",' + num_days + ',1,Holidays!$A:$A)'
    }
    else {
        var formula = '=WORKDAY.INTL("' + formatDate + '",' + num_days + ',1,Holidays!$A:$A)'
    }

    // var formula = 
    //  '=IF(B'+program+'="MERN"  ,WORKDAY.INTL(D'+program+',59,1,Holidays!$A:$A),WORKDAY.INTL(D'+program+',99,1,Holidays!$A:$A))'
    //  tempCalcCell.setFormula(formula)
    tempCalcCell.setValue(formula)
    var mltnDate = tempCalcCell.getValue()
    return Utilities.formatDate(new Date(mltnDate), "GMT+1", "MM/dd/yyyy")


}

function addMilestoneEvents(emailList, eventProgram, eventCohort1, date) {
    var programSheet = getProgramSheet(eventProgram)
    var startRow = 2;
    var startCol = 1;
    var lastDay = programSheet.getLastRow();
    var programRange = programSheet.getRange(startRow, startCol, lastDay, 6)
    var programList = programRange.getValues()
    var delta = 0

    for (var i = 0; i < programList.length; i++) {
        var programDay = programList[i]
        var milestone = programDay[5]
        if (milestone != "") {
            var workDays = programDay[0]



            //   var offDays = 0

            var emailListObj = []
            for (var j = 0; j < emailList.length; j++) {
                emailListObj.push({ email: emailList[j] })
            }


            //    delta += offDays
            var milestoneDate = getMilestoneDate(date, workDays, eventProgram)
            //   var milestoneDate = addDays(date, offDays + workDays)

            var title = "Milestone-" + eventProgram + '-' + eventCohort1
            var startDate = new Date(milestoneDate).setHours(9)
            var endDate = new Date(milestoneDate).setHours(18)
            var event = {
                summary: title,
                description: milestone,
                start: {
                    dateTime: new Date(startDate).toISOString()
                },
                end: {
                    dateTime: new Date(endDate).toISOString()
                },
                attendees: emailListObj,
                colorId: 11,
                transparency: 'transparent'
            };
            event = Calendar.Events.insert(event, 'primary')
            //    
            //    var event = CalendarApp.createAllDayEvent(title, milestoneDate,  {
            //    guests	: emailList,
            //    sendInvites: false,
            //    description: milestone,
            //  
            //    });
            //    event.setColor('11');
            //  
            //     event.transparency = 'transparent'
        }
    }
}

function addDays(date, days) {
    var result = new Date(date);
    result.setDate(result.getDate() + days);

    return result;
}


function addSpikeEvent(emailList, name, description, startDate, endDuration, location) {

    var event = CalendarApp.getDefaultCalendar().createEvent(name,
        startDate,
        endDuration,
        {
            location: location,
            guests: emailList,
            sendInvites: true,
            description: description
        });
}

function addStandUpEvent(emailList, name, description, startDate, endDuration, location, endDate) {
    var endMonth = new Date(startDate)
    endMonth.setMonth(startDate.getMonth() + 1)

    var title = 'Stand Up ' + name
    var eventSeries = CalendarApp.getDefaultCalendar().createEventSeries(title,
        startDate,
        endDuration,
        CalendarApp.newRecurrence().addWeeklyRule()
            .onlyOnWeekdays([CalendarApp.Weekday.MONDAY, CalendarApp.Weekday.TUESDAY, CalendarApp.Weekday.WEDNESDAY, CalendarApp.Weekday.THURSDAY, CalendarApp.Weekday.FRIDAY])
            .until(endDate),
        {
            location: location,
            guests: emailList,
            sendInvites: true,
            description: description
        });
}


function getStaff(program) {
    var staffSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Staff")
    var lastStaff = staffSheet.getLastRow();
    var staffRange = staffSheet.getRange(2, 1, lastStaff, 3)
    var staffList = staffRange.getValues()
    var emailList = []
    for (var i = 0; i < staffList.length; i++) {

        if (staffList[i][0] == program || program == "ALL") {
            emailList.push(staffList[i][2])
        }
        else if (staffList[i][0] == "FULL-STACK" && program == "JAVA" || staffList[i][0] == "FULL-STACK" && program == "MERN") {
            emailList.push(staffList[i][2])
        }
    }

    return emailList
}


function nextweek(today) {

    var nextweek = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 7);
    return nextweek;
}