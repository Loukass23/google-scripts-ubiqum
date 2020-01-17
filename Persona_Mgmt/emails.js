function mainController() {
    var configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config")
    var program = configSheet.getRange("B1").getValue()
    var location = configSheet.getRange("B2").getValue()
    var url = configSheet.getRange("B3").getValue()
    var ss = SpreadsheetApp.openByUrl(url);


    //loop through all Ubiqum mgmt sheets
    var relationSheet = ss.getSheetByName("Ubiqum")
    var startRow = 2;
    var startCol = 1
    var lastCourseLocation = relationSheet.getLastRow() - 1;
    var relationRange = relationSheet.getRange(startRow, startCol, lastCourseLocation, 3);
    var urlList = relationRange.getValues();
    for (i in urlList) {
        var url = urlList[i][2]
        var locationUbiqum = urlList[i][0]
        if (url !== "") {
            if (location == locationUbiqum || location == "ALL") {
                var ss = SpreadsheetApp.openByUrl(url);
                studentsController(ss, program, locationUbiqum)
            }
        }
    }
}

function studentsController(ss, program, location) {
    var studentSheet = ss.getSheetByName("Students")
    var emailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Emails")
    var hourNow = new Date().getHours()

    var startRow = 2;
    var startCol = 1
    var lastEmail = emailSheet.getLastRow() - 1;
    var emailRange = emailSheet.getRange(startRow, startCol, lastEmail, 11);

    var lastStudent = studentSheet.getLastRow() - 1;
    var studentRange = studentSheet.getRange(2, 2, lastStudent, 9)

    var studentsList = studentRange.getValues();
    var emailList = emailRange.getValues();

    for (i in studentsList) {
        var student = studentsList[i];
        if (program == student[0] || program == 'ALL' || (program == 'FULL-STACK' && (student[0] == 'MERN' || student[0] == 'JAVA'))) {
            emailController(student, emailList, hourNow, location)
        }
    }
}

function emailController(student, emailList, hourNow, location) {
    var emailAddress = student[3]
    var studentName = student[2]
    var program = student[0]

    for (j in emailList) {
        var email = emailList[j]
        var subject = email[7]

        if (isRecipient(student, email, hourNow) && !alreadySent(emailAddress, subject)) {
            writeLogs(emailAddress, subject, location, program)
            sendEmail(email, emailAddress, studentName)

        }
    }
}
function greetingPicker(type, name, hour) {
    if (type == "Informal") {
        var hello = ["Hey", "Hi", "Dear", "Hello"]
        var index = Math.floor(Math.random() * Math.floor(hello.length));
        return hello[index] + " " + name + ","
    }
    else if (type == "Formal") {
        if (hour < 12) return "Good morning " + name + ","
        else if (hour < 17) return "Good afternoon " + name + ","
        else return "Good evening " + name + ","
    }
    else return ""
}


function sendEmail(email, emailAddress, studentName) {
    var signature = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config").getRange("B6").getValue()

    var attachment = email[10]
    var hour = email[6]
    var subject = email[7]
    var greeting = email[8]
    var body = email[9]


    //check for img signature
    if (signature) {
        var signatureBlob = DriveApp.getFileById(signature).getBlob();
        var signature = "<img src = 'cid:signature' />  ";
    }
    else var signature = "";

    var message = '<p>' + greetingPicker(greeting, studentName, hour) + '</p>' + body + signature

    //case attachment + signature
    if (attachment != "" && signature) {
        var file = DriveApp.getFileById(attachment);

        MailApp.sendEmail({
            to: emailAddress,
            subject: subject,
            attachments: [file],
            htmlBody: message,
            inlineImages:
            {
                signature: signatureBlob
            }
        });
    }
    //case attachment + no signature
    else if (attachment != "" && !signature) {
        var file = DriveApp.getFileById(attachment);

        MailApp.sendEmail({
            to: emailAddress,
            subject: subject,
            attachments: [file],
            htmlBody: message,
        });
    }
    //case no attachment / signature
    else if (attachment == "" && signature) MailApp.sendEmail({
        to: emailAddress,
        subject: subject,
        htmlBody: message,
        inlineImages:
        {
            signature: signatureBlob
        }
    });
    else MailApp.sendEmail({
        to: emailAddress,
        subject: subject,
        htmlBody: message,

    });

}

function isRecipient(student, email, hourNow) {
    var type = email[0]
    var moduleMail = email[1]

    var programStudent = student[0]
    var moduleStdent = student[4]
    var sprintStdent = student[5]
    var dayStdent = student[6]
    var refDayStdent = student[7]

    var sprintMail = email[2]
    var dayMail = email[3]
    var refDay = email[4]
    var hourMail = email[6]
    var programtMail = email[5]


    if (type == "Sprint" && moduleStdent == moduleMail) {
        if (programStudent == programtMail || programtMail == "ALL" || (programtMail == 'FULL-STACK' && (programStudent == 'MERN' || programStudent == 'JAVA'))) {
            //passed or current Sprint and Day
            if ((sprintMail == sprintStdent && dayMail < dayStdent)
                ||
                (sprintMail == sprintStdent && dayMail == dayStdent && hourNow >= hourMail)
                ||
                (sprintMail < sprintStdent)) {
                return true
            }
        }
    }
    else if (type == "Ref Day") {

        //check student module
        if (programStudent == programtMail || programtMail == "ALL" || (programtMail == 'FULL-STACK' && (programStudent == 'MERN' || programStudent == 'JAVA'))) {

            //check ref days
            if (refDay < refDayStdent || refDay == refDayStdent && hourNow >= hourMail) {
                return true
            }
        }

    }
    else return false

}

