var spreadsheetId = '1E82wrm8FP9MftoKQkWQAK7ecz9LhoJdSSS2OiPzisGc';
var tutors = [];
var daysOfTheWeek = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday'];
// names of the columns that represent a day of the week.
var columns = ['B','C','D','E','F','G'];
var allShifts = ['9-10','10-11','11-12','12-1','1-2','2-3','3-4','4-5','5-6','6-7','7-8','8-9'];
// ranges of all shifts in which tutoring is not offered.
var nonActiveShifts = ['B2:B25','B42:B49','G22:G49'];

var SURVEY_NAME = 'Form Responses 1';
var SCHEDULE_SHEET = 'New Schedule';
var MAX_TUTORS = 4;
var DAYS_WITH_INDV_AND_GROUP = 4;
var STARTING_ROW = 2;
var FRIDAY_HOURS_CELL = 4;
var SUNDAY_HOURS_CELL = 9;
var LAST_SHIFT = allShifts.indexOf('1-2');

/**
 * Creates a menu entry in the Google Sheets UI when the document is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Custom Menu')
    .addItem('Create schedule', 'main')
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

function doGet() {
  return HtmlService.createHtmlOutputFromFile('sidebar');
}

/**
 * Gets the range currently selected by the user.
 */
function findWaitlistForSelection() {
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(SCHEDULE_SHEET);
  var range = SpreadsheetApp.getActive().getActiveRange().getA1Notation();
  fetchSurveyData();
  var allShiftRanges = getAllShiftRanges();
  // check if selected range fits the valid ranges of shifts.
  var isValidRange = validateRange(range, allShiftRanges);
  if (isValidRange) {
    // return list of tutors that are waitlisted for the selected range.
    var day = getDayFromRange(range);
    var shift = getShiftFromRange(range);
    var waitlist = getWaitlistedTutors(range, day, shift);
    var waitlistNames = waitlist.map(function(tutor) { return tutor.name });
    return waitlistNames;
  } else {
    throw new Error('Invalid range selected.');
  }
}

/**
 * Given a range in A1 notation, return whether it is within the bounds
 * of a shift on the schedule.
 */
function validateRange(range, allValidRanges) {
  for (var i = 0; i < allValidRanges.length; i++) {
    if (range === allValidRanges[i]) return true;
    // TODO: add check if in non active shift    
  }
  return false;
}

function getDayFromRange(range) {
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(SCHEDULE_SHEET);
  var column = sheet.getRange(range).getColumn();
  return sheet.getRange(1, column).getValue().toLowerCase();
}

function getShiftFromRange(range) {
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(SCHEDULE_SHEET);
  var row = sheet.getRange(range).getRow();
  return sheet.getRange(row, 1).getValue();
}

function getWaitlistedTutors(range, day, shift) {
  var currentTutors = getTutorsOnScheudle(range);
  Logger.log(currentTutors);
  var waitlist = tutors.filter(function(tutor) {
    var notOnShift = currentTutors.indexOf(tutor.name) === -1;
    var canWorkShift = tutor.shifts[day].indexOf(shift) >= 0;
    return (notOnShift && canWorkShift);
  });
  return waitlist;
}

function getTutorsOnScheudle(range) {
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(SCHEDULE_SHEET);
  // getValues returns a 2d array, indexed by row then column
  var onSchedule = sheet.getRange(range).getValues();
  // flatten result into a single 1d array for simpler iteration
  var currentTutors = [].concat.apply([], onSchedule);
  // remove all empty data points
  currentTutors = currentTutors.filter(function(tutor) {
    return tutor !== '';
  });
  return currentTutors;
}

/**
 * Returns an array containing ranges in A1 format, each of which represent 
 * a shift block for the given schedule.
 */
function getAllShiftRanges() {
  var allShiftRanges = [];
  for (var i = 0; i < daysOfTheWeek.length; i++) {
    var column = columns[i];
    var startRow = STARTING_ROW;
    var endRow = startRow + MAX_TUTORS-1; // subtract one because the group of cells is inclusive
    for (var j = 0; j < allShifts.length; j++) {
      var cluster = column+startRow + ':' + column+endRow;
      // store the range of the current shift block
      allShiftRanges.push(cluster);
      startRow += MAX_TUTORS;
      endRow += MAX_TUTORS;
    }
  }
  return allShiftRanges;
}

function main() {
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  try {
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName(SCHEDULE_SHEET));
  } catch(e) {
    // insert after form responses
    spreadsheet.insertSheet(SCHEDULE_SHEET, 1);
  }
  writeBlankSchedule(SCHEDULE_SHEET);
  fetchSurveyData();
  tutors = sortByGivenHours_(tutors);
  // construct schedule
  for (var i = 0; i < daysOfTheWeek.length; i++) {
    var currentDayTutors = createSchedule(tutors, daysOfTheWeek[i]);
    var shiftRow = STARTING_ROW;
    for (var j = 0; j < currentDayTutors.length; j++) {
      var dayColumn = columns[i];
      var currentShift = currentDayTutors[j];
      writeToSchedule(currentShift, dayColumn, shiftRow);
      shiftRow += MAX_TUTORS;
    }  
  }
  // ensure that there is no invalid data
  clearNonActiveShifts();
}

/**
 * Creates a Sheets API service object and constructs tutor objects based on 
 * survey responses.
 */
function fetchSurveyData() {
  tutors = []; // reset data
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

function sortByGivenHours_(tutors) {
  tutors.sort(function(a,b) {
    return a.givenHours - b.givenHours;
  });
  return tutors; // bc apparently js doesn't pass by reference >:(
}

/** 
 * Pre-formats a blank schedule with the days of the week, shift hours, and 
 * an empty grid. Also populates the array for all shift ranges.
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
  var shiftRow = STARTING_ROW;
  allShifts.forEach(function(shift) {
    sheet.getRange(leftColumn+shiftRow)
      .setValue(shift)
      .setFontWeight('bold')
      .setHorizontalAlignment('right');
    shiftRow += MAX_TUTORS;  
  });
  // per shift per day
  var shiftBlocks = getAllShiftRanges();
  shiftBlocks.forEach(function(block) {
    sheet.getRange(block)
      .setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  });
}

function clearNonActiveShifts() {
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(SCHEDULE_SHEET);
  nonActiveShifts.forEach(function(range) {
    // block out shift on sheet
    sheet.getRange(range).clear().setBackground('#D3D3D3');
  }); 
}

/**
 * Returns an array of arrays - each of which represents a shift,
 * and contains the tutors that work that shift.
 */
function createSchedule(tutors, dayOfWeek) {
  var schedule = [];
  for (var i = 0; i < allShifts.length; i++) {
    schedule[i] = [];
    // no more active shifts after Friday, 1-2 pm 
    if (dayOfWeek === 'friday' && i > LAST_SHIFT) {
      break;
    }
    tutors.forEach(function(tutor) {
      if (tutor.shifts[dayOfWeek].indexOf(allShifts[i]) != -1) {
        schedule[i].push(tutor); 
      }
    });
  }
  return schedule;
}

/**
 * Populates the range of cells for a given day and shift, 
 * represented by the column and row respectively, with the names
 * of the tutors that work during that shift.
 */
function writeToSchedule(tutorsPerShift, column, row) { 
  // TODO: add waitlisting for non-priority tutors
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(SCHEDULE_SHEET);
  var cell = sheet.getRange(column+row);
  for (var count = 0; count < MAX_TUTORS; count++) {
    if (!tutorsPerShift[count]) {
      cell.setValue('');
    } else {
      // write the tutor's name in the shift's cell
      cell.setValue(tutorsPerShift[count].name);
    }
    row++;
    cell = sheet.getRange(column+row);
  }
}