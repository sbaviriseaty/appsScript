/**
 * Change these to match the column names you are using for email 
 * recepient addresses and email sent column.
 */
const RECIPIENT_COL = "Email";
const EMAIL_SENT_COL = "Delivery";

var NUM_ITEMS_TO_RANK = 5;
var ACTIVITIES_PER_PERSON = 1;
var NUM_TEST_USERS = 250;

/**
 * Add custom menu items when opening the sheet.
 */
function onOpen() {
    var menu = SpreadsheetApp.getUi().createMenu('Events')
        .addItem('Link form', 'buildForm_')
        .addItem('Generate test data', 'generateTestData_')
        .addItem('Assign events', 'assignEvents_')
        .addItem('Send assignments', 'sendEmails')
        .addToUi();
}

/**
 * Builds a form based on the "Activity Schedule" sheet. The form asks attendees to rank their top
 * N choices of activities, where N is defined by NUM_ITEMS_TO_RANK.
 */
function buildForm_() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var form = FormApp.openByUrl('https://docs.google.com/forms/d/1sTOtBPtU9rUGB-InEYh-pL9GPUnVkB2jpMqGjvJrsco/edit')
}

/**
 * Assigns activities using a random priority/random serial dictatorship approach. The results
 * are then populated into two new sheets, one listing activities per person, the other listing
 * the rosters for each activity.
 *
 * See https://en.wikipedia.org/wiki/Random_serial_dictatorship for additional information.
 */
function assignEvents_() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var activities = loadActivitySchedule_(ss);
    var activityIds = activities.map(function(activity) {
        return activity.id;
    });
    var attendees = loadAttendeeResponses_(ss, activityIds);
    assignWithRandomPriority_(attendees, activities, 1);
    writeAttendeeAssignments_(ss, attendees);
    writeActivityRosters_(ss, activities);
}

/**
 * Select activities via random priority.
 *
 * @param {object[]} attendees - Array of attendees to assign activities to
 * @param {object[]} activities - Array of all available activities
 * @param {number} numActivitiesPerPerson - Maximum number of activities to assign
 */
function assignWithRandomPriority_(attendees, activities, numActivitiesPerPerson) {
    var activitiesById = activities.reduce(function(obj, activity) {
        obj[activity.id] = activity;
        return obj;
    }, {});
    for (var i = 0; i < numActivitiesPerPerson; ++i) {
        var randomizedAttendees = shuffleArray_(attendees);
        randomizedAttendees.forEach(function(attendee) {
            if (attendee.type != 'General') {
                makeChoice_(attendee, activitiesById);
            }
        });
    }
}

/**
 * Attempt to assign an activity for an attendee based on their preferences and current schedule.
 *
 * @param {object} attendee - Attendee looking to join an activity
 * @param {object} activitiesById - Map of all available activities
 */
function makeChoice_(attendee, activitiesById) {

    for (var i = 0; i < attendee.preferences.length; ++i) {
        var activity = activitiesById[attendee.preferences[i]];
        //if (attendee.time == 'Shift 1+2') {
        //  continue;
        //}
        if (!activity) {
            continue;
        }
        var canJoin = checkAvailability_(attendee, activity);
        if (canJoin) {
            attendee.assigned.push(activity);
            activity.roster.push(attendee);
            break;
        }
    }
}

/**
 * Checks that an activity has capacity and doesn't conflict with previously assigned
 * activities.
 *
 * @param {object} attendee - Attendee looking to join the activity
 * @param {object} activity - Proposed activity
 * @return {boolean} - True if attendee can join the activity
 */
function checkAvailability_(attendee, activity) {
    if (activity.capacity <= activity.roster.length) {
        return false;
    }
    //console.log(activity.time);
    if (activity.time != attendee.time)
    {
      return false;
    }

    var timesConflict = attendee.assigned.some(function(assignedActivity) {
        return !(assignedActivity.startAt.getTime() > activity.endAt.getTime() ||
            activity.startAt.getTime() > assignedActivity.endAt.getTime());
    });

    //if (attendee.time == 'Shift 1+2') {
    //    timesConflict = true;
    //}
    return !timesConflict;
};

/**
 * Populates a sheet with the assigned activities for each attendee.
 *
 * @param {Spreadsheet} ss - Spreadsheet to write to.
 * @param {object[]} attendees - Array of attendees with their activity assignments
 */
function writeAttendeeAssignments_(ss, attendees) {
    var sheet = findOrCreateSheetByName_(ss, 'Events by person');
    sheet.clear();
    sheet.appendRow(['Name', 'Email', 'Type', 'Time', 'Assignment']);
    var rows = attendees.map(function(attendee) {
        // Prefill row to ensure consistent length otherwise
        // can't bulk update the sheet with range.setValues()
        var row = fillArray_([], 5, '');
        row[0] = attendee.name;
        row[1] = attendee.email;
        row[2] = attendee.type;
        row[3] = attendee.time;
        //console.log(attendee.time);
        attendee.assigned.forEach(function(activity, index) {
            row[4] = activity.description;
        });
        return row;
    });
    bulkAppendRows_(sheet, rows);
    sheet.setFrozenRows(1);
    sheet.getRange('1:1').setFontWeight('bold');
    sheet.autoResizeColumns(1, sheet.getLastColumn());
}

/**
 * Populates a sheet with the rosters for each activity.
 *
 * @param {Spreadsheet} ss - Spreadsheet to write to.
 * @param {object[]} activities - Array of activities with their rosters
 */
function writeActivityRosters_(ss, activities) {
    var sheet = findOrCreateSheetByName_(ss, 'Event rosters');
    sheet.clear();
    var rows = [];
    var rows = activities.map(function(activity) {
        var roster = activity.roster.map(function(attendee) {
            return attendee.email;
        });
        return [activity.description].concat(roster).concat(roster.length);        
    });
    // Transpose the data so each activity is a column
    rows = transpose_(rows, '');
    bulkAppendRows_(sheet, rows);
    sheet.setFrozenRows(1);
    sheet.getRange('1:1').setFontWeight('bold');
    sheet.autoResizeColumns(1, sheet.getLastColumn());
}

/**
 * Loads the activity schedule.
 *
 * @param {Spreadsheet} ss - Spreadsheet to load from
 * @return {object[]} Array of available activities.
 */
function loadActivitySchedule_(ss) {
    var timeZone = ss.getSpreadsheetTimeZone();
    var sheet = ss.getSheetByName('Event Schedule');
    var rows = sheet.getSheetValues(
        sheet.getFrozenRows() + 1, 1,
        sheet.getLastRow() - 1, sheet.getLastRow());
    var activities = rows.map(function(row, index) {
        var time;
        var name = row[0];
        var startAt = new Date(row[1]);
        var endAt = new Date(row[2]);
        var capacity = parseInt(row[3]);
        var formattedStartAt = Utilities.formatDate(startAt, timeZone, 'hh:mm a');
        //console.log(formattedStartAt);
        var formattedEndAt = Utilities.formatDate(endAt, timeZone, 'hh:mm a');
        var description = Utilities.formatString('%s (%s-%s)', name, formattedStartAt, formattedEndAt);
        if (formattedStartAt == "07:30 AM" && formattedEndAt == "04:00 PM") {
          time = "Shift 1+2";
        } else if (formattedStartAt == "07:30 AM" && formattedEndAt == "12:30 PM") {
          time = "Shift 1";
        } else {
          time = "Shift 2";
        }
        return {
            id: index,
            name: name,
            description: description,
            capacity: capacity,
            startAt: formattedStartAt,
            endAt: formattedEndAt,
            roster: [],
            time: time,
        };
    });
    //console.log(activities)
    return activities;
}

/**
 * Loads the attendeee response data
 *
 * @param {Spreadsheet} ss - Spreadsheet to load from
 * @param {number[]} allActivityIds - Full set of available activity IDs
 * @return {object[]} Array of parsed attendee respones.
 */
function loadAttendeeResponses_(ss, allActivityIds) {
    var sheet = findResponseSheetForForm_(ss);

    if (!sheet || sheet.getLastRow() == 1) {
        return undefined;
    }

    var rows = sheet.getSheetValues(
        sheet.getFrozenRows() + 1, 1,
        sheet.getLastRow() - 1, sheet.getLastRow());

    var attendees = rows.map(function(row) {
        var _ = row.shift(); // Ignore timestamp
        var email = row.shift();
        var first_name = row.shift();
        var last_name = row.shift();
        var phone_number = row.shift();
        var DOB = row.shift();
        var background_check = row.shift();
        var student = row.shift();
        var UFID = row.shift();
        var college = row.shift();
        var major = row.shift();
        var year = row.shift();
        var type = row.shift();
        var timeG = row.shift();
        var timeE = row.shift();
        var autoAssign = 'Yes';
        // Find ranked items in the response data.
        var preferences = row.reduce(function(prefs, value, index) {
            var match = value.match(/(\d+).*/);
            if (timeG == 'Shift 1+2' || timeE == 'Shift 1+2') {}
            if (type == '') {
                type = 'General Volunteer';
            }
            if (type == 'General Volunteer') {
                autoAssign = 'No';
                timeE = timeG;
                return prefs;
            }
            if (!match || type == 'General Volunteer') {
                return prefs;
            }
            var rank = parseInt(match[1]) - 1; // Convert ordinal to array index
            prefs[rank] = index;
            return prefs;
        }, []);
        if (autoAssign == 'Yes') {
            // If auto assigning additional activites, append a randomized list of all the activities.
            // These will then be considered as if the attendee ranked them.
            var additionalChoices = shuffleArray_(allActivityIds);
            preferences = preferences.concat(additionalChoices);
        }
        return {
            email: email,
            preferences: preferences,
            assigned: [],
            type: type,
            time: timeE,
            name: first_name + " " + last_name
        };
    });
    //console.log(attendees);
    return attendees;
}

/**
 * Simulates a large number of users responding to the form. This enables users to quickly
 * experience the full solution without having to collect sufficient form responses
 * through other means.
 */
function generateTestData_() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = findResponseSheetForForm_(ss);
    if (!sheet) {
        var msg = 'No response sheet found. Create the form and try again.';
        SpreadsheetApp.getUi().alert(msg);
    }
    if (sheet.getLastRow() > 1) {
        var msg = 'Response sheet is not empty, can not generate test data. ' +
            'Remove responses and try again.';
        SpreadsheetApp.getUi().alert(msg);
        return;
    }
    var activities = loadActivitySchedule_(ss);
    var extra = fillArray_([], 9, '');
    var empty = fillArray_([], 86, '');
    var choices = fillArray_([], 22, '');
    range_(1, 5).forEach(function(value) {
        choices[value] = toOrdinal_(value);
    });

    var rows = range_(1, NUM_TEST_USERS).map(function(value) {
        var randomizedChoices = shuffleArray_(choices);
        var email = Utilities.formatString('person%d@example.com', value);
        var name = Utilities.formatString('person %d', value);
        var random = Math.floor(Math.random() * 100);
        if (random <= 75) {
            if (random <= 25) {
                return [new Date(), email, name].concat(extra).concat(['Event']).concat(['']).concat(['Shift 2']).concat(empty).concat(randomizedChoices);
            } else if (random <= 50) {
                return [new Date(), email, name].concat(extra).concat(['Event']).concat(['']).concat(['Shift 1']).concat(randomizedChoices).concat(empty);
            } else {
                both = fillArray_([], activities.length, '');
                range_(1, 5).forEach(function(value) {
                    both[value] = toOrdinal_(value);
                });
                var randomizedBoth = shuffleArray_(both);
                return [new Date(), email, name].concat(extra).concat(['Event']).concat(['']).concat(['Shift 1+2']).concat(randomizedBoth);
            }
        } else {
            randomizedChoices = fillArray_([], activities.length, '');
            //console.log(activities.length);
            //console.log("General");
            if (random <= 83) {
                return [new Date(), email, name].concat(extra).concat(['General']).concat(['Shift 1']).concat(['']).concat(randomizedChoices);
            } else if (random <= 91) {
                return [new Date(), email, name].concat(extra).concat(['General']).concat(['Shift 2']).concat(['']).concat(randomizedChoices);
            } else {
                return [new Date(), email, name].concat(extra).concat(['General']).concat(['Shift 1+2']).concat(['']).concat(randomizedChoices);
            }
        }
    });
    bulkAppendRows_(sheet, rows);
}

/**
 * Retrieve a sheet by name, creating it if it doesn't yet exist.
 *
 * @param {Spreadsheet} ss - Containing spreadsheet
 * @Param {string} name - Name of sheet to return
 * @return {Sheet} Sheet instance
 */
function findOrCreateSheetByName_(ss, name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) {
        return sheet;
    }
    return ss.insertSheet(name);
}

/**
 * Faster version of appending multiple rows via ranges. Requires all rows are equal length.
 *
 * @param {Sheet} sheet - Sheet to append to
 * @param {Array<Array<object>>} rows - Rows to append
 */
function bulkAppendRows_(sheet, rows) {
    var startRow = sheet.getLastRow() + 1;
    var startColumn = 1;
    var numRows = rows.length;
    var numColumns = rows[0].length;
    sheet.getRange(startRow, startColumn, numRows, numColumns).setValues(rows);
}

/**
 * Copies and randomizes an array
 *
 * @param {object[]} array - Array to shuffle
 * @return {object[]} randomized copy of the array
 */
function shuffleArray_(array) {
    var clone = array.slice(0);
    for (var i = clone.length - 1; i > 0; i--) {
        var j = Math.floor(Math.random() * (i + 1));
        var temp = clone[i];
        clone[i] = clone[j];
        clone[j] = temp;
    }
    return clone;
}

/**
 * Formats an number as an ordinal.
 *
 * See: https://stackoverflow.com/questions/13627308/add-st-nd-rd-and-th-ordinal-suffix-to-a-number/13627586
 *
 * @param {number} i - Number to format
 * @return {string} Formatted string
 */
function toOrdinal_(i) {
    var j = i % 10;
    var k = i % 100;
    if (j == 1 && k != 11) {
        return i + 'st';
    }
    if (j == 2 && k != 12) {
        return i + 'nd';
    }
    if (j == 3 && k != 13) {
        return i + 'rd';
    }
    return i + 'th';
}

/**
 * Locates the sheet containing the form responses.
 *
 * @param {Spreadsheet} ss - Spreadsheet instance to search
 * @return {Sheet} Sheet with form responses, undefined if not found.
 */
function findResponseSheetForForm_(ss) {
    var formUrl = ss.getFormUrl();
    if (!ss || !formUrl) {
        return undefined;
    }
    var sheets = ss.getSheets();
    for (var i in sheets) {
        if (sheets[i].getFormUrl() === formUrl) {
            return sheets[i];
        }
    }
    return undefined;
}

/**
 * Fills an array with a value ([].fill() not supported in Apps Script.)
 *
 * @param {object[]} arr - Array to fill
 * @param {number} length - Number of items to fill.
 * @param {object} value - Value to place at each index.
 * @return {object[]} the array, for chaining purposes
 */
function fillArray_(arr, length, value) {
    for (var i = 0; i < length; ++i) {
        arr[i] = value;
    }
    return arr;
}

/**
 * Creates an fills an array with numbers in the range [start, end].
 *
 * @param {number} start - First value in the range, inclusive
 * @param {number} end - Last value in the range, inclusive
 * @return {number[]} Array of values representing the range
 */
function range_(start, end) {
    var arr = [start];
    var i = start;
    while (i < end) {
        arr.push(i += 1);
    }
    return arr;
}

/**
 * Transposes a matrix/2d array. For cases where the rows are not the same length,
 * `fillValue` is used where no other value would otherwise be present.
 *
 * @param {Array<Array<object>>} arr - 2D array to transpose
 * @param {object} fillValue - Placeholder for undefined values created as a result
 *     of the transpose. Only required if rows aren't all of equal length.
 * @return {Array<Array<object>>} New transposed array
 */
function transpose_(arr, fillValue) {
    var transposed = [];
    arr.forEach(function(row, rowIndex) {
        row.forEach(function(col, colIndex) {
            transposed[colIndex] = transposed[colIndex] || fillArray_([], arr.length, fillValue);
            transposed[colIndex][rowIndex] = row[colIndex];
        });
    });
    return transposed;
}

/**
 * Iterates row by row in the input range and returns an array of objects.
 * Each object contains all the data for a given row, indexed by its normalized column name.
 * @param {Sheet} sheet The sheet object that contains the data to be processed
 * @param {Range} range The exact range of cells where the data is stored
 * @param {number} columnHeadersRowIndex Specifies the row number where the column names are stored.
 *   This argument is optional and it defaults to the row immediately above range;
 * @return {object[]} An array of objects.
 */
function getRowsData(sheet, range, columnHeadersRowIndex) {
    columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
    var numColumns = range.getEndColumn() - range.getColumn() + 1;
    var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
    var headers = headersRange.getValues()[0];
    return getObjects(range.getValues(), normalizeHeaders(headers));
}

/**
 * For every row of data in data, generates an object that contains the data. Names of
 * object fields are defined in keys.
 * @param {object} data JavaScript 2d array
 * @param {object} keys Array of Strings that define the property names for the objects to create
 * @return {object[]} A list of objects.
 */
function getObjects(data, keys) {
    var objects = [];
    for (var i = 0; i < data.length; ++i) {
        var object = {};
        var hasData = false;
        for (var j = 0; j < data[i].length; ++j) {
            var cellData = data[i][j];
            if (isCellEmpty(cellData)) {
                continue;
            }
            object[keys[j]] = cellData;
            hasData = true;
        }
        if (hasData) {
            objects.push(object);
        }
    }
    return objects;
}

/**
 * Returns an array of normalized Strings.
 * @param {string[]} headers Array of strings to normalize
 * @return {string[]} An array of normalized strings.
 */
function normalizeHeaders(headers) {
    var keys = [];
    for (var i = 0; i < headers.length; ++i) {
        var key = normalizeHeader(headers[i]);
        if (key.length > 0) {
            keys.push(key);
        }
    }
    return keys;
}

/**
 * Normalizes a string, by removing all alphanumeric characters and using mixed case
 * to separate words. The output will always start with a lower case letter.
 * This function is designed to produce JavaScript object property names.
 * @param {string} header The header to normalize.
 * @return {string} The normalized header.
 * @example "First Name" -> "firstName"
 * @example "Market Cap (millions) -> "marketCapMillions
 * @example "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
 */
function normalizeHeader(header) {
    var key = '';
    var upperCase = false;
    for (var i = 0; i < header.length; ++i) {
        var letter = header[i];
        if (letter == ' ' && key.length > 0) {
            upperCase = true;
            continue;
        }
        if (!isAlnum(letter)) {
            continue;
        }
        if (key.length == 0 && isDigit(letter)) {
            continue; // first character must be a letter
        }
        if (upperCase) {
            upperCase = false;
            key += letter.toUpperCase();
        } else {
            key += letter.toLowerCase();
        }
    }
    return key;
}

/**
 * Returns true if the cell where cellData was read from is empty.
 * @param {string} cellData Cell data
 * @return {boolean} True if the cell is empty.
 */
function isCellEmpty(cellData) {
    return typeof(cellData) == 'string' && cellData == '';
}

/**
 * Returns true if the character char is alphabetical, false otherwise.
 * @param {string} char The character.
 * @return {boolean} True if the char is a number.
 */
function isAlnum(char) {
    return char >= 'A' && char <= 'Z' ||
        char >= 'a' && char <= 'z' ||
        isDigit(char);
}

/**
 * Returns true if the character char is a digit, false otherwise.
 * @param {string} char The character.
 * @return {boolean} True if the char is a digit.
 */
function isDigit(char) {
    return char >= '0' && char <= '9';
}

/**
 * Sends emails from spreadsheet rows.
 */
function sendEmails() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dataSheet = ss.getSheets()[0];
    var dataRange = dataSheet.getRange(2, 1, dataSheet.getMaxRows() - 1, 4);

    var templateSheet = ss.getSheets()[1];
    var emailTemplate = templateSheet.getRange('F2').getValue();

    // Create one JavaScript object per row of data.
    var objects = getRowsData(dataSheet, dataRange);

    // For every row object, create a personalized email from a template and send
    // it to the appropriate person.
    for (var i = 0; i < objects.length; ++i) {
        // Get a row object
        var rowData = objects[i];

        // Generate a personalized email.
        // Given a template string, replace markers (for instance ${"First Name"}) with
        // the corresponding value in a row object (for instance rowData.firstName).
        var emailText = fillInTemplateFromObject(emailTemplate, rowData);
        var emailSubject = 'UFSO Volunteer Assignment';

        MailApp.sendEmail(rowData.emailAddress, emailSubject, emailText);
    }
}

/**
 * Replaces markers in a template string with values define in a JavaScript data object.
 * @param {string} template Contains markers, for instance ${"Column name"}
 * @param {object} data values to that will replace markers.
 *   For instance data.columnName will replace marker ${"Column name"}
 * @return {string} A string without markers. If no data is found to replace a marker,
 *   it is simply removed.
 */
function fillInTemplateFromObject(template, data) {
    var email = template;
    // Search for all the variables to be replaced, for instance ${"Column name"}
    var templateVars = template.match(/\$\{\"[^\"]+\"\}/g);

    // Replace variables from the template with the actual values from the data object.
    // If no value is available, replace with the empty string.
    for (var i = 0; templateVars && i < templateVars.length; ++i) {
        // normalizeHeader ignores ${"} so we can call it directly here.
        var variableData = data[normalizeHeader(templateVars[i])];
        email = email.replace(templateVars[i], variableData || '');
    }

    return email;
}