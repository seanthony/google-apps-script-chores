import 'google-apps-script';


function launch() {
    var spreadSheetId = '??';
    var documentId = '??';
    var calendarId = '??';
    var stringDate = "December 2, 2019 00:00:00 CST";
    assignChores(stringDate, spreadSheetId, documentId, calendarId);
}


function assignChores(stringDate, spreadSheetId, documentId, calendarId) {
    var ss = SpreadsheetApp.openById(spreadSheetId);
    var choreObject = makeChores(ss.getSheetByName('Chore Description').getDataRange().getValues());

    var sheet = ss.getSheetByName('Chore_List');
    var choreData = sheet.getDataRange().getValues();

    while (true) {
        var roster = shuffle(ss.getSheetByName('Roster').getDataRange().getValues().slice(1));
        for (var c = 1; c < choreData.length; c++) {
            choreData[c][6] = roster.pop()[0];

            if (roster.length < 1) {
                roster = shuffle(ss.getSheetByName('Roster').getDataRange().getValues().slice(1)).slice();
            }
        }

        if (isChoreDataValid(choreData)) {
            break;
        }
    }

    var choreData = addDates(choreData, stringDate);
    sheet.getDataRange().setValues(choreData);

    makeWeeklyPrintOut(choreData, documentId);

    var roster = ss.getSheetByName('Roster').getDataRange().getValues();
    var emailObject = makeEmailObject(roster);

    // assign chores
    for (var c = 1; c < choreData.length; c++) {
        var choreTitle = choreData[c][3];
        var startTime = Utilities.formatDate(new Date(choreData[c][5]), 'US/Central', 'MMMM dd, yyyy HH:mm:ss z');
        var choreDescription = choreObject[choreData[c][2]]['description'];
        var email = emailObject[choreData[c][6]];
        var eventId = sendInvite(choreTitle, startTime, choreDescription, email, calendarId);
        choreData[c][7] = eventId;
    }

    sheet.getDataRange().setValues(choreData);
}

function shuffle(array) {
    var currentIndex = array.length, temporaryValue, randomIndex;

    // While there remain elements to shuffle...
    while (0 !== currentIndex) {
        // Pick a remaining element...
        randomIndex = Math.floor(Math.random() * currentIndex);
        currentIndex -= 1;
        // And swap it with the current element.
        temporaryValue = array[currentIndex];
        array[currentIndex] = array[randomIndex];
        array[randomIndex] = temporaryValue;
    }

    return array;
}


function makeChores(values) {
    var header = values[0].slice();
    var choreObject = {};
    for (var c = 1; c < values.length; c++) {
        var key = values[c][0];
        choreObject[key] = {};
        for (var i = 1; i < header.length; i++) {
            var innerkey = header[i];
            var val = values[c][i];
            choreObject[key][innerkey] = val;
        }
    }
    return choreObject;
}

function isChoreDataValid(choreData) {
    var seen = {};
    var day = '';
    for (var c = 1; c < choreData.length; c++) {
        var newDay = choreData[c][0];
        var name = choreData[c][6];
        if (newDay != day) {
            day = newDay;
            seen = {
                name: true
            };
        } else {
            if (seen[choreData[c][6]]) {
                return false;
            } else {
                seen[choreData[c][6]] = true;
            }
        }
    }

    return true;

}

function addDates(choreData, stringDate) {
    var mondayDate = new Date(stringDate);
    var dateObj = {
        'Monday': mondayDate,
        'Tuesday': new Date(mondayDate.valueOf() + 1000 * 60 * 60 * 24 * 1),
        'Wednesday': new Date(mondayDate.valueOf() + 1000 * 60 * 60 * 24 * 2),
        'Thursday': new Date(mondayDate.valueOf() + 1000 * 60 * 60 * 24 * 3),
        'Friday': new Date(mondayDate.valueOf() + 1000 * 60 * 60 * 24 * 4)
    }


    for (var c = 1; c < choreData.length; c++) {
        var weekday = choreData[c][0];
        var date = dateObj[weekday];
        var day = Utilities.formatDate(date, 'US/Central', 'MM/dd/yyyy');
        var time = Utilities.formatDate(new Date(choreData[c][4]), 'US/Central', 'hh:mm a');
        var strDate = day + ' ' + time + ' CDT';
        var calDate = Utilities.formatDate(new Date(strDate), 'US/Central', 'MM/dd/yyyy hh:mm a');
        choreData[c][1] = day;
        choreData[c][5] = calDate;
        choreData[c][7] = '';
    }

    return choreData;
}

function makeWeeklyPrintOut(choreData, documentId) {
    var doc = DocumentApp.openById(documentId);
    var body = doc.getBody();
    body.clear();

    var tableStarter = ['Day', 'Chore', 'Name', 'Signature'];

    var cells = [
        tableStarter.slice(),
        [choreData[1][0], choreData[1][3], choreData[1][6], '                            ']
    ];

    for (var c = 2; c < choreData.length; c++) {
        if (choreData[c][0] != choreData[c - 1][0]) {
            body.appendTable(cells);
            body.appendPageBreak();
            var cells = [
                tableStarter.slice()
            ];
        }
        cells.push([choreData[c][0], choreData[c][3], choreData[c][6], '                            ']);
    }

    body.appendTable(cells);
}


function makeEmailObject(roster) {
    var emails = {};
    for (var c = 1; c < roster.length; c++) {
        emails[roster[c][0]] = roster[c][1];
    }

    return emails;
}


function sendInvite(choreTitle, startTime, choreDescription, email, calendarId) {
    var calendar = CalendarApp.getCalendarById(calendarId);

    var date = new Date(startTime);
    var endDate = new Date(date.valueOf() + 15 * 60 * 1000);
    var endTime = Utilities.formatDate(endDate, 'US/Central', 'MMMM dd, yyyy HH:mm:ss z');

    var event = calendar.createEvent(choreTitle,
        new Date(startTime),
        new Date(endTime),
        {
            location: 'Base Camp',
            description: choreDescription,
            guests: email,
            sendInvites: true
        }
    );

    return event.getId();
}
