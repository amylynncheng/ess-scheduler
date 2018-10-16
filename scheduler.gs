var spreadsheetId = '1E82wrm8FP9MftoKQkWQAK7ecz9LhoJdSSS2OiPzisGc';
var scheduleSheet = 'New Schedule';
var tutors = [];
var daysOfTheWeek = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday'];
var columns = ['B','C','D','E','F','G'];
var allShifts = ['9-10','10-11','11-12','12-1','1-2','2-3','3-4','4-5','5-6','6-7','7-8','8-9'];
var nonActiveShifts = ['B2:B25','B42:B49','G22:G49'];

var SURVEY_NAME = 'Form Responses 1';
var MAX_TUTORS = 4;
var DAYS_WITH_INDV_AND_GROUP = 4;
var STARTING_CELL = 2;
var FRIDAY_HOURS_CELL = 4;
var SUNDAY_HOURS_CELL = 9;

/**
 * Creates a menu entry in the Google Sheets UI when the document is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Custom Menu')
    .addItem('Show sidebar', 'showSidebar')
    .addToUi();
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 */
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Scheduling Sidebar');
  SpreadsheetApp.getUi().showSidebar(html);
}

function main() {
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  try {
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName(scheduleSheet));
  } catch(e) {
    // insert after form responses
    spreadsheet.insertSheet(scheduleSheet, 1);
  }
  writeBlankSchedule(scheduleSheet);
}

/**
 * Creates a Sheets API service object and constructs tutor objects based on 
 * survey responses.
 */
function fetchSurveyData() {
  const survey = SpreadsheetApp.openById(spreadsheetId).getSheetByName(SURVEY_NAME);
  const lastRow = survey.getLastRow();
  const infoRange = survey.getRange('A2:F' + lastRow);
  const basicInfo = infoRange.getValues();
  const timeRange = survey.getRange('G2:P' + lastRow);
  const hoursInfo = timeRange.getValues();
  if (!basicInfo) {
    Logger.log('No data found.');
  } else {
    for (var row = 0; row < basicInfo.length; row++) {
      var tutor = new Object();
      tutor.row = row;
      tutor.timestamp = basicInfo[row][0] // A
      tutor.name = basicInfo[row][1]; // B
      tutor.email = basicInfo[row][2] // C
      tutor.major = basicInfo[row][3]; // D
      tutor.level = basicInfo[row][4]; // E
      tutor.courses = basicInfo[row][5]; // F
      
      // combine individual and group hours.
      var totalHours = [];
      // only required for Monday to Thursday.
      for (var day = 0; day < DAYS_WITH_INDV_AND_GROUP; day++) {
        var individual = hoursInfo[row][day];
        // individual and group responses for the same day are 5 cells apart.
        var group = hoursInfo[row][day+5];
        totalHours.push(mergeHours_(individual, group));
      }
      // add the hours for Friday and Sunday as they are.
      totalHours.push(hoursInfo[row][FRIDAY_HOURS_CELL]);
      totalHours.splice(0, 0, hoursInfo[row][SUNDAY_HOURS_CELL]); // push to front
      tutor.shifts = getHours_(totalHours);
      
      // optional: total hours tutor is willing to work
      tutor.givenHours = getGivenHours(tutor);
      tutors.push(tutor);
    }
  }
}

/**
 * Combines the responses for individual and group hours per day into
 * one single array. 
 * 
 * @param {array} Available hours per work day from a tutor's form response.
 */
function mergeHours_(individual, group) {
  if (!individual && !group) { // both are empty
    return '';
  } else if (!group) { // only works individual hours
    return individual;
  } else if (!individual) { // only works group hours
    return group;
  }
  return individual + ', ' + group;
}

/**
 * Converts the responses given for hours available per day into an object
 * containing each day of the the week --> an array of shifts the tutor
 * can work for that day. If the tutor cannot work at all for a given day,
 * the value for that day is empty.
 *
 * @param {array} Available hours per work day from a tutor's form response.
 * @return {object} {workday1: [hours], workday2: [hours], ...}
 */
function getHours_(hoursPerDay) {
  var shifts = new Object();
  for (var day = 0; day < hoursPerDay.length; day++) {
    if (!hoursPerDay[day]) { // tutor does not work this day.
      shifts[daysOfTheWeek[day]] = [];
    } else {
      shifts[daysOfTheWeek[day]] = formatHours_(hoursPerDay[day]);
    }
  }
  return shifts;
}

/**
 * Removes the meridian suffix from each shift duration in the input.
 *
 * @param {string} shifts in the format "9-10 AM, 10-11 AM,..."
 * @return {array} shifts in the format [9-10, 10-11,...]
 */
function formatHours_(hours) {
  var array = hours.split(', ');
  for (var i = 0; i < array.length; i++) {
    var noSuffix = array[i].split(' ')[0];
    array[i] = noSuffix;
  }
  return array;
}

/**
 * Returns the total number of hours the given tutor is ABLE to work.
 * @param tutor
 */
function getGivenHours(tutor) {
  var hours = 0;  
  for (var day = 0; day < daysOfTheWeek.length; day++) {
    var shiftsPerDay = tutor.shifts[daysOfTheWeek[day]];
    if (shiftsPerDay[0] === '') { // tutor does not work this day
      continue;
    }
    // each shift is one hour, so we can just add the length of the array containing all shifts
    hours += shiftsPerDay.length;
  }
  return hours;
}

/** 
 * Pre-formats a blank schedule with the days of the week, shift hours, and 
 * an empty grid. 
 */
function writeBlankSchedule(sheetName) {
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  sheet.clear();
  // top row
  sheet.getRange('B1:G1')
    .setValues([['Sunday','Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']])
    .setFontWeight('bold')
  // left column
  var leftColumn = 'A';
  var shiftRow = STARTING_CELL;
  allShifts.forEach(function(shift) {
    sheet.getRange(leftColumn+shiftRow)
      .setValue(shift)
      .setFontWeight('bold');
    shiftRow += MAX_TUTORS;  
  });
  // per shift per day
  for (var i = 0; i < daysOfTheWeek.length; i++) {
    var column = columns[i];
    var startRow = STARTING_CELL;
    var endRow = startRow + MAX_TUTORS-1; // subtract one because the group of cells is inclusive
    for (var j = 0; j < allShifts.length; j++) {
      var cluster = column+startRow + ':' + column+endRow;
      sheet.getRange(cluster)
        .setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
      startRow += MAX_TUTORS;
      endRow += MAX_TUTORS;
    }
  }
  // block out non-active shifts (Sunday and Friday)
  nonActiveShifts.forEach(function(range) {
    blockOut_(sheet, range);
  });  
}

function blockOut_(sheet, range) {
  sheet.getRange(range).setBackground('#D3D3D3');
}