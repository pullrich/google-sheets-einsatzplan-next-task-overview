function generateNextTaskOverview() {
    // https://findicons.com/files/icons/2212/carpelinx/64/add.png
    removeOverviewSheet();


    var dateColumn = 1;
    var todayDate = new Date();
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var dataSheet = activeSpreadsheet.getSheetByName(getDataSheetName());
    if (dataSheet != null) {
        Logger.log('Found data sheet by name.');
    } else {
        return;
    }

    const LAST_NAME_COLUMN_INDEX = 18;
    const NAMES_ROW_INDEX = 2;

    var allValues = dataSheet.getDataRange().getValues();

    var dateColumnArray = getColumnDataFromRowColumnArray(allValues, 0);
    var firstDateOccurrsIndex = getIndexOfDateValueOccurrance(dateColumnArray, false);
    var lastDateOccurrsIndex = getIndexOfDateValueOccurrance(dateColumnArray, true);

    var taskRowIndex = -1;
    var dateOfTasks = null;
    for (var i = firstDateOccurrsIndex; i < lastDateOccurrsIndex + 1; i++) {
        if (isValidDate(dateColumnArray[i])) {
            if (isGreaterOrEqualDate(dateColumnArray[i], todayDate)) {
                taskRowIndex = i;
                dateOfTasks = dateColumnArray[i];
                break;
            }
        }
        continue;
    }

    if (taskRowIndex === -1) {
        Logger.log('No fitting date row found.');
        return;
    }

    var assignmentList = [];
    for (var i = 1; i < LAST_NAME_COLUMN_INDEX + 1; i++) {
        assignmentList.push({ task: allValues[taskRowIndex][i], agent: allValues[NAMES_ROW_INDEX][i] })
    }
    Logger.log('assignment list:');
    Logger.log(assignmentList);

    var testarray = ['c', 'b', 'a'];
    Logger.log('testarray:');
    Logger.log(testarray);
    testarray.sort();
    Logger.log('sorted testarray:');
    Logger.log(testarray);

    assignmentList.sort(function (a, b) { (a.task > b.task) ? 1 : -1 });
    //assignmentList.sort();
    //assignmentList.reverse();
    Logger.log('Sorted assignment list:');
    Logger.log(assignmentList);
    // TODO: Sorting does not seem to work ... at least it is not displayed as expected in the log.

    //return;

    var userVisibleSheet = activeSpreadsheet.getActiveSheet();
    var activeRange = userVisibleSheet.getActiveRange();

    var overviewSheet = activeSpreadsheet.insertSheet('Aufgabenübersicht');

    activeSpreadsheet.setActiveSheet(userVisibleSheet);
    activeSpreadsheet.setActiveRange(activeRange);

    addOverviewToSheet(overviewSheet, 1, 1, dateOfTasks, assignmentList);
}

function addOverviewToSheet(sheet, row, column, date, assignmentList) {
    var headingRange = sheet.getRange(row, column, 1, 2);
    headingRange.merge();
    headingRange.setValue('Heading');

    for (var i = 0; i < assignmentList.length; i++) {
        sheet.getRange(row + i + 1, column).setValue(assignmentList[i].task);
        sheet.getRange(row + i + 1, column + 1).setValue(assignmentList[i].agent);
    }
    sheet.autoResizeColumn(1);
    sheet.autoResizeColumn(2);

    headingRange.setValue('Übersicht für den: ' + Utilities.formatDate(date, getTimeZoneGermany(), "dd.MM.yyyy"));
    headingRange.setFontSize(12);
    headingRange.setFontWeight('bold');
}

function getColumnDataFromRowColumnArray(anArray, columnIndex) {
    var newArray = [];
    for (var i = 0; i < anArray.length; i++) {
        newArray.push(anArray[i][columnIndex]);
    }
    return newArray;
}


function removeOverviewSheet() {
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var overviewSheet = activeSpreadsheet.getSheetByName("Aufgabenübersicht");
    if (overviewSheet != null) {
        activeSpreadsheet.deleteSheet(overviewSheet);
    } else {
        return;
    }
}

function getIndexOfDateValueOccurrance(dates, backwards) {
    if (backwards) {
        for (var i = dates.length - 1; 0 < i; i--) {
            var value = dates[i];
            if (isValidDate(value)) {
                return i;
            }
        }
        return null;
    }
    else {
        for (var i = 0; i < dates.length - 1; i++) {
            var value = dates[i];
            if (isValidDate(value)) {
                return i;
            }
        }
        return null;
    }
}

function getTimeZoneGermany() {
    return "GMT+2";
}

function getTodayDate() {
    return new Date();
}

function getNameRow() {
    return 3;
}

function getDataSheetName() {
    return "Tabellenblatt1";
}

function isGreaterOrEqualDate(d1, d2) {
    return (Utilities.formatDate(d1, getTimeZoneGermany(), "yyyy-MM-dd") >= Utilities.formatDate(d2, getTimeZoneGermany(), "yyyy-MM-dd"));
}

function getAllValues() {
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var dataSheet = activeSpreadsheet.getSheetByName(getDataSheetName());
    if (dataSheet === null) {
        Logger.log('Could not find sheet by name: ' + getDataSheetName());
        return null;
    }
    return dataSheet.getDataRange().getValues();
}

function getRowOfRelevantTasks(allValues, date) {
    var dateColumnArray = getColumnDataFromRowColumnArray(allValues, 0);
    var firstDateOccurrsIndex = getIndexOfDateValueOccurrance(dateColumnArray, false);
    var lastDateOccurrsIndex = getIndexOfDateValueOccurrance(dateColumnArray, true);

    var taskRowIndex = -1;
    var dateOfTasks = null;
    for (var i = firstDateOccurrsIndex; i < lastDateOccurrsIndex + 1; i++) {
        if (isValidDate(dateColumnArray[i])) {
            if (isGreaterOrEqualDate(dateColumnArray[i], date)) {
                taskRowIndex = i;
                dateOfTasks = dateColumnArray[i];
                break;
            }
        }
        continue;
    }

    return taskRowIndex + 1;
}

function isChangeInRelevantRow(changedRow) {
    return (changedRow === getRowOfRelevantTasks(getAllValues(), getTodayDate()) || changedRow === getNameRow());
}

/**
 * The event handler triggered when opening the spreadsheet.
 * @param {Event} e The onOpen event.
 */
function onOpen(e) {
    //generateNextTaskOverview();
}

/**
 * The event handler triggered when editing the spreadsheet.
 * @param {Event} e The onEdit event.
 */
function onEdit(e) {
    if (isChangeInRelevantRow(e.range.getRow())) {
        generateNextTaskOverview();
    }
    else {
        Logger.log('Change not in relevant row. No action required.');
    }
}

// From http://stackoverflow.com/questions/1353684
// Returns 'true' if variable d is a date object.
function isValidDate(d) {
    if (Object.prototype.toString.call(d) !== "[object Date]")
        return false;
    return !isNaN(d.getTime());
}

// Test if value is a date and if so format
// otherwise, reflect input variable back as-is. 
function isDate(sDate) {
    if (isValidDate(sDate)) {
        sDate = Utilities.formatDate(new Date(sDate), "PST", "MM/dd/yyyy");
    }
    return sDate;
}