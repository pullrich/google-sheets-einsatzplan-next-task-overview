function generateNextTaskOverview() {
    generateTaskOverview(getTodayDate());
}

function generateTaskOverview(date) {
    removeOverviewSheet();

    const LAST_NAME_COLUMN_INDEX = 18;
    const NAMES_ROW_INDEX = 2;

    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var allValues = getAllValues();

    var taskRowIndex = getRowOfRelevantTasks(allValues, date) - 1;
    if (taskRowIndex <= -1) {
        Logger.log('No fitting date row found.');
        return;
    }
    var dateOfTasks = getDateOfTasks(allValues, taskRowIndex + 1);

    var assignmentList = [];
    for (var i = 1; i < LAST_NAME_COLUMN_INDEX + 1; i++) {
        assignmentList.push({ task: allValues[taskRowIndex][i], agent: allValues[NAMES_ROW_INDEX][i] });
    }

    assignmentList.sort(compareTaskAssignmentsByTask);

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

function getDateOfTasks(allValues, row) {
    return allValues[row - 1][0];
}

function isGreaterOrEqualDate(d1, d2) {
    return (Utilities.formatDate(d1, getTimeZoneGermany(), "yyyy-MM-dd") >= Utilities.formatDate(d2, getTimeZoneGermany(), "yyyy-MM-dd"));
}

function compareTaskAssignmentsByTask(a, b) {
    if (a.task < b.task) {
        return -1;
    }
    if (a.task > b.task) {
        return 1;
    }
    return 0;
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
    generateNextTaskOverview();
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