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
	deleteSheet('Übersicht-h');
	deleteSheet('Übersicht-v');

	const LAST_NAME_COLUMN_INDEX = 18;
	const NAMES_ROW_INDEX = 2;

	var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
	var sourceSheet = activeSpreadsheet.getSheetByName(getDataSheetName());
	var allValues = getAllValues();

	var taskRowIndex = getRowOfRelevantTasks(allValues, date) - 1;
	if (taskRowIndex <= -1) {
		Logger.log('No fitting date row found.');
		return;
	}
	var dateOfTasks = getDateOfTasks(allValues, taskRowIndex + 1);

	// Fill array with overview data.
	var assignmentList = [];
	for (var i = 1; i < LAST_NAME_COLUMN_INDEX + 1; i++) {
		var tasks = allValues[taskRowIndex][i].split('/').map(function (e) { return e.trim(); });
		var agent = allValues[NAMES_ROW_INDEX][i];

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