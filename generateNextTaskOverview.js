function generateNextTaskOverview() {
  // https://findicons.com/files/icons/2212/carpelinx/64/add.png
  removeOverviewSheet();


  var dateColumn = 1;
  var todayDate = new Date();
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = activeSpreadsheet.getSheetByName("Tabellenblatt1");
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

  layoutOverviewOnSheet(overviewSheet, 1, 1, dateOfTasks, assignmentList);
}

function layoutOverviewOnSheet(sheet, row, column, date, assignmentList) {
  var headingRange = sheet.getRange(row, column, 1, 2);
  headingRange.merge();
  headingRange.setValue('Heading');

  for (var i = 0; i < assignmentList.length; i++) {
    sheet.getRange(row + i + 1, column).setValue(assignmentList[i].task);
    sheet.getRange(row + i + 1, column + 1).setValue(assignmentList[i].agent);
  }
  sheet.autoResizeColumn(1);
  sheet.autoResizeColumn(2);

  headingRange.setValue('Übersicht für den: ' + Utilities.formatDate(date, "GMT+1", "dd.MM.yyyy"));
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

function isGreaterOrEqualDate(d1, d2) {
  return (Utilities.formatDate(d1, "GMT+1", "yyyy-MM-dd") >= Utilities.formatDate(d2, "GMT+1", "yyyy-MM-dd"));
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
  // Set a comment on the edited cell to indicate when it was changed.
  //var range = e.range;
  //range.setNote('Last modified: ' + new Date());
  generateNextTaskOverview();
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