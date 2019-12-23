/**
 * @OnlyCurrentDoc
 */
// https://developers.google.com/apps-script/guides/services/authorization

/**
 * The event handler triggered when opening the spreadsheet.
 * @param {Event} e The onOpen event.
 */
function onOpen(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  menuEntries.push({ name: "Nächste", functionName: "generateNextTaskOverview" });
  //menuEntries.push(null); // line separator
  menuEntries.push({ name: "Für Datum", functionName: "genOverview_showDatePrompt" });

  ss.addMenu("Übersicht erzeugen", menuEntries);


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

function TaskAssignment(task, agent, taskcolor) {
  this.task = task;
  this.agent = agent.trim();
  this.taskcolor = taskcolor;

  var namePieces = this.agent.split(' ');
  this.lastname = namePieces[namePieces.length - 1];
  this.firstname = namePieces[0];

  this.getLastnameFirstname = function () {
    return this.lastname + ', ' + this.firstname;
  }
}



function generateNextTaskOverview() {
  generateTaskOverview(getTodayDate());
}

function generateTaskOverview(date) {
  // The user might interfere - so there may be old intermediate sheets around which need to be deleted.
  deleteSheet('...erstelle Übersicht-h...');
  deleteSheet('...erstelle Übersicht-v...');

  deleteSheet(getOverviewSheetName_h());
  deleteSheet(getOverviewSheetName_v());

  var ui = SpreadsheetApp.getUi();
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = activeSpreadsheet.getSheetByName(getDataSheetName());
  var allValues = getAllValues();
  var namesRowIndex = getNameRow(allValues) - 1;
  var lastNameColumnIndex = getLastNameColumn(allValues) - 1;

  var taskRowIndex = getRowOfRelevantTasks(allValues, date) - 1;
  if (taskRowIndex <= -1) {
    ui.alert('Da keine geeignete Datenzeile gefunden werden konnte, wird kein Übersichtstabellenblatt erzeugt.', ui.ButtonSet.OK);
    Logger.log('No fitting date row found.');
    return;
  }
  var dateOfTasks = getDateOfTasks(allValues, taskRowIndex + 1);

  // Fill array with overview data.
  var assignmentList = [];
  for (var i = 1; i < lastNameColumnIndex + 1; i++) {
    var tasks = allValues[taskRowIndex][i].split('/').map(function (e) { return e.trim(); });
    var agent = allValues[namesRowIndex][i];

    tasks.forEach(function (aTask) {
      assignmentList.push(new TaskAssignment(aTask, agent, sourceSheet.getRange(taskRowIndex + 1, i + 1).getBackground()));
    });
  }

  writeOverviewToSheetVertically(dateOfTasks, assignmentList);
  writeOverviewToSheetHorizontally(dateOfTasks, assignmentList);
}

function writeOverviewToSheetHorizontally(date, assignments) {
  sheet = addHiddenSheet('...erstelle Übersicht-h...');
  assignments.sort(compareTaskAssignmentsByTask);

  var row = 1;
  var column = 1;
  var firstHeaderRow = row;
  var firstHeaderColumn = column;
  row += getHeadingRowCount();

  for (var i = 0; i < assignments.length; i++) {
    sheet.getRange(row, column).setValue(assignments[i].task);
    sheet.getRange(row, column).setBackground(assignments[i].taskcolor);
    sheet.getRange(row, column + 1).setValue(assignments[i].getLastnameFirstname());
    row += 1;
  }

  row = 1;
  column = 4;
  var secondHeaderRow = row;
  var secondHeaderColumn = column;

  assignments.sort(compareTaskAssignmentsByLastnameFirstname);

  row += getHeadingRowCount();
  for (var i = 0; i < assignments.length; i++) {
    sheet.getRange(row, column).setValue(assignments[i].getLastnameFirstname());
    sheet.getRange(row, column + 1).setValue(assignments[i].task);
    sheet.getRange(row, column + 1).setBackground(assignments[i].taskcolor);
    row += 1;
  }

  sheet.autoResizeColumn(1);
  sheet.autoResizeColumn(2);
  sheet.setColumnWidth(3, 10);
  sheet.autoResizeColumn(4);
  sheet.autoResizeColumn(5);

  writeHeading(sheet, firstHeaderRow, firstHeaderColumn, 'Übersicht für den ' + getDateInGermanFormat(date), 'nach Aufgaben');
  writeHeading(sheet, secondHeaderRow, secondHeaderColumn, 'Übersicht für den ' + getDateInGermanFormat(date), 'nach Namen');


  sheet.setName(getOverviewSheetName_h());
  sheet.showSheet();
}

function writeOverviewToSheetVertically(date, assignments) {
  sheet = addHiddenSheet('...erstelle Übersicht-v...');
  assignments.sort(compareTaskAssignmentsByTask);

  var row = 1;
  var column = 1;
  var firstHeaderRow = row;
  var firstHeaderColumn = column;
  row += getHeadingRowCount();

  for (var i = 0; i < assignments.length; i++) {
    sheet.getRange(row, column).setValue(assignments[i].task);
    sheet.getRange(row, column).setBackground(assignments[i].taskcolor);
    sheet.getRange(row, column + 1).setValue(assignments[i].getLastnameFirstname());
    row += 1;
  }

  row += 1;
  var secondHeaderRow = row;
  var secondHeaderColumn = column;
  row += getHeadingRowCount();

  assignments.sort(compareTaskAssignmentsByLastnameFirstname);

  for (var i = 0; i < assignments.length; i++) {
    sheet.getRange(row, column).setValue(assignments[i].getLastnameFirstname());
    sheet.getRange(row, column + 1).setValue(assignments[i].task);
    sheet.getRange(row, column + 1).setBackground(assignments[i].taskcolor);
    row += 1;
  }

  sheet.autoResizeColumn(1);
  sheet.autoResizeColumn(2);

  writeHeading(sheet, firstHeaderRow, firstHeaderColumn, 'Übersicht für den ' + getDateInGermanFormat(date), 'nach Aufgaben');
  writeHeading(sheet, secondHeaderRow, secondHeaderColumn, 'Übersicht für den ' + getDateInGermanFormat(date), 'nach Namen');

  sheet.setName(getOverviewSheetName_v());
  sheet.showSheet();
}

function writeHeading(sheet, row, column, text, sortHintText) {
  firstLineRange = sheet.getRange(row, column, 1, 2);
  firstLineRange.merge();
  firstLineRange.setValue(text);
  firstLineRange.setFontSize(11);

  secondLineRange = sheet.getRange(row + 1, column, 1, 2);
  secondLineRange.merge();
  secondLineRange.setValue(sortHintText);
  secondLineRange.setFontSize(9);

  headingRange = sheet.getRange(row, column, 2, 2);
  headingRange.setFontWeight('bold');
  headingRange.setBackground('silver');
}

function getHeadingRowCount() {
  return 2;
}

function addHiddenSheet(sheetname) {
  var userVisibleSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var activeRange = userVisibleSheet.getActiveRange();

  var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetname);
  sheet.hideSheet();

  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(userVisibleSheet);
  SpreadsheetApp.getActiveSpreadsheet().setActiveRange(activeRange);

  return sheet;
}

function getColumnDataFromRowColumnArray(anArray, columnIndex) {
  var newArray = [];
  for (var i = 0; i < anArray.length; i++) {
    newArray.push(anArray[i][columnIndex]);
  }
  return newArray;
}

function deleteSheet(sheetname) {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetToDelete = activeSpreadsheet.getSheetByName(sheetname);
  if (sheetToDelete != null) {
    activeSpreadsheet.deleteSheet(sheetToDelete);
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

function getNameRow(allValues) {
  // Expectiation: The name-row will be the row before the first occurrance of a date value in the first column.
  const FIRST_COLUMN_INDEX = 0;
  var dateColumnArray = getColumnDataFromRowColumnArray(allValues, FIRST_COLUMN_INDEX);
  var firstDateOccurrsIndex = getIndexOfDateValueOccurrance(dateColumnArray, false);
  var nameRow = firstDateOccurrsIndex; // We can just use the index, because we need the row (and not the index) BEFORE the first occurrence of a date.
  return nameRow;
}

function getLastNameColumn(allValues) {
  // Expectation: The first column will not contain a name in the name row. The following columns will contain names if the column is not empty.
  //    So the first empty column in the name row marks the end of the names.
  var lastNameColumn = 0;
  var nameRowIndex = getNameRow(allValues) - 1;
  for (var i = 1; i < allValues[nameRowIndex].length; i++) {
    var cellValue = allValues[nameRowIndex][i];
    if (cellValue === "" || cellValue === undefined || cellValue === null) {
      lastNameColumn = i; // We can just use the index, because we need the colun (and not the index) BEFORE the first occurrence of an empty cell.
      break;
    }
  }
  return lastNameColumn;
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

function compareTaskAssignmentsByLastnameFirstname(a, b) {
  if (a.getLastnameFirstname() < b.getLastnameFirstname()) {
    return -1;
  }
  if (a.getLastnameFirstname() > b.getLastnameFirstname()) {
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
  return (changedRow === getRowOfRelevantTasks(getAllValues(), getTodayDate()) || changedRow === getNameRow(getAllValues()));
}

/**
 * Expected format DD.MM.YYYY
 */
function parseTextAsDate(text) {
  const DATE_ELEMENTS_EXP = 3;
  const DAY_IDX = 0;
  const MONTH_IDX = 1;
  const YEAR_IDX = 2;

  var dateElements = text.split(".");

  var year = parseInt(dateElements[YEAR_IDX]);
  var month = parseInt(dateElements[MONTH_IDX]);
  var day = parseInt(dateElements[DAY_IDX]);

  // IMPROVE: Use return object { ok: true|false, date: date, error: text } ??
  if (dateElements.length < DATE_ELEMENTS_EXP) {
    throw new Error("Not enough date elements.");
  }
  if (isNaN(year) || isNaN(month) || isNaN(day)) {
    throw new Error("A date element is NaN.");
  }

  return new Date(year, month - 1, day);
}

function genOverview_showDatePrompt() {
  var ui = SpreadsheetApp.getUi();

  var result = ui.prompt(
    'Übersicht für ein Datum erzeugen',
    'Bitte gib das gewünschte Datum im Format ' + getDateFormatString() + ' ein.',
    ui.ButtonSet.OK_CANCEL);

  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK) {
    dateFromUser = new Date(0);
    try {
      dateFromUser = parseTextAsDate(text);
    }
    catch (err) {
      Logger.log(err);
      ui.alert('Eingegebenes Datum [' + text + '] nicht erkannt.\nBenötigtes Format: ' + getDateFormatString());
      return;
    }

    generateTaskOverview(dateFromUser);
  } else if (button == ui.Button.CANCEL) {
  } else if (button == ui.Button.CLOSE) {
  }
}

//
// Helper functions ->
//

function getDateFormatString() {
  return "DD.MM.YYYY"
}

function getDateInGermanFormat(date) {
  return Utilities.formatDate(date, getTimeZoneGermany(), "dd.MM.yyyy")
}

// From http://stackoverflow.com/questions/1353684
// Returns 'true' if variable d is a date object.
function isValidDate(d) {
  if (Object.prototype.toString.call(d) !== "[object Date]")
    return false;
  return !isNaN(d.getTime());
}

function getOverviewSheetName_h() {
  return "Übersicht-h";
}

function getOverviewSheetName_v() {
  return "Übersicht-v";
}

function getTimeZoneGermany() {
  return "GMT+2";
}

function getTodayDate() {
  return new Date();
}
