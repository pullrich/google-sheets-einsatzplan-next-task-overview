/**
 * @OnlyCurrentDoc
 */
// https://developers.google.com/apps-script/guides/services/authorization

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

  deleteSheet('Übersicht-h');
  deleteSheet('Übersicht-v');

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

  var firstHeadingRange = sheet.getRange(1, 1, 1, 2);
  firstHeadingRange.merge();
  firstHeadingRange.setValue('Heading');
  firstHeadingRange.setBackground('silver');

  assignments.sort(compareTaskAssignmentsByTask);

  var row = 2;
  var column = 1;
  for (var i = 0; i < assignments.length; i++) {
    sheet.getRange(row, column).setValue(assignments[i].task);
    sheet.getRange(row, column).setBackground(assignments[i].taskcolor);
    sheet.getRange(row, column + 1).setValue(assignments[i].getLastnameFirstname());
    row += 1;
  }

  column = 4;
  var secondHeadingRange = sheet.getRange(1, column, 1, 2);
  secondHeadingRange.merge();
  secondHeadingRange.setValue('Heading');
  secondHeadingRange.setBackground('silver');

  assignments.sort(compareTaskAssignmentsByLastnameFirstname);

  row = 2;
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

  firstHeadingRange.setValue('Übersicht für den ' + Utilities.formatDate(date, getTimeZoneGermany(), "dd.MM.yyyy") + ' nach Aufgaben');
  firstHeadingRange.setFontSize(11);
  firstHeadingRange.setFontWeight('bold');

  secondHeadingRange.setValue('Übersicht für den ' + Utilities.formatDate(date, getTimeZoneGermany(), "dd.MM.yyyy") + ' nach Namen');
  secondHeadingRange.setFontSize(11);
  secondHeadingRange.setFontWeight('bold');

  sheet.setName('Übersicht-h');
  sheet.showSheet();
}

function writeOverviewToSheetVertically(date, assignments) {
  sheet = addHiddenSheet('...erstelle Übersicht-v...');

  var firstHeadingRange = sheet.getRange(1, 1, 1, 2);
  firstHeadingRange.merge();
  firstHeadingRange.setValue('Heading');
  firstHeadingRange.setBackground('silver');

  assignments.sort(compareTaskAssignmentsByTask);

  var row = 2;
  var column = 1;
  for (var i = 0; i < assignments.length; i++) {
    sheet.getRange(row, column).setValue(assignments[i].task);
    sheet.getRange(row, column).setBackground(assignments[i].taskcolor);
    sheet.getRange(row, column + 1).setValue(assignments[i].getLastnameFirstname());
    row += 1;
  }

  row += 1;
  var secondHeadingRange = sheet.getRange(row, 1, 1, 2);
  secondHeadingRange.merge();
  secondHeadingRange.setValue('Heading');
  secondHeadingRange.setBackground('silver');
  row += 1;

  assignments.sort(compareTaskAssignmentsByLastnameFirstname);

  for (var i = 0; i < assignments.length; i++) {
    sheet.getRange(row, column).setValue(assignments[i].getLastnameFirstname());
    sheet.getRange(row, column + 1).setValue(assignments[i].task);
    sheet.getRange(row, column + 1).setBackground(assignments[i].taskcolor);
    row += 1;
  }

  sheet.autoResizeColumn(1);
  sheet.autoResizeColumn(2);

  firstHeadingRange.setValue('Übersicht für den ' + Utilities.formatDate(date, getTimeZoneGermany(), "dd.MM.yyyy") + ' nach Aufgaben');
  firstHeadingRange.setFontSize(11);
  firstHeadingRange.setFontWeight('bold');

  secondHeadingRange.setValue('Übersicht für den ' + Utilities.formatDate(date, getTimeZoneGermany(), "dd.MM.yyyy") + ' nach Namen');
  secondHeadingRange.setFontSize(11);
  secondHeadingRange.setFontWeight('bold');

  sheet.setName('Übersicht-v');
  sheet.showSheet();
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

function getTimeZoneGermany() {
  return "GMT+2";
}

function getTodayDate() {
  return new Date();
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
 * The event handler triggered when opening the spreadsheet.
 * @param {Event} e The onOpen event.
 */
function onOpen(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  // When the user clicks on "addMenuExample" then "Menu Entry 1", the function function1 is
  // executed.
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

// From http://stackoverflow.com/questions/1353684
// Returns 'true' if variable d is a date object.
function isValidDate(d) {
  if (Object.prototype.toString.call(d) !== "[object Date]")
    return false;
  return !isNaN(d.getTime());
}


/**
 * Expected format DD.MM.YYYY
 */
function parseTextAsDate(text) {
  var dayIdx = 0;
  var monthIdx = 1;
  var yearIdx = 2;
  var dateElements = text.split(".");
  return new Date(parseInt(dateElements[yearIdx]), parseInt(dateElements[monthIdx]) - 1, parseInt(dateElements[dayIdx]));
}

function genOverview_showDatePrompt() {
  var ui = SpreadsheetApp.getUi();

  var result = ui.prompt(
    'Übersicht für ein Datum erzeugen',
    'Bitte gib das gewünschte Datum im Format ' + getDateFormatString() + ' ein.',
    ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK) {
    dateFromUser = new Date(0);
    try {
      dateFromUser = parseTextAsDate(text);
    }
    catch (err) {
      ui.alert('Eingegebenes Datum [' + text + '] nicht erkannt.\nBenötigtes Format: ' + getDateFormatString());
      return;
    }

    //formattedDateFromUser = Utilities.formatDate(dateFromUser, getTimeZoneGermany(), "dd.MM.yyyy")
    //ui.alert('Eigegebenes Datum: ' + formattedDateFromUser + '\nErzeuge Übersicht.');
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
