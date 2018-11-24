//-------------------------- CONSTANTS --------------------------
// Names
var SURVEY_NAME = 'Form Responses 1';
var SCHEDULE_SHEET = 'New Schedule';

// Schedule properties
var daysOfTheWeek = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday'];
// names of the columns that represent a day of the week.
var ALL_SHIFTS = ['9-10','10-11','11-12','12-1','1-2','2-3','3-4','4-5','5-6','6-7','7-8','8-9'];
var MAX_TUTORS = 4; 
var DAYS_WITH_INDV_AND_GROUP = 4;
var FRIDAY_HOURS_CELL = 4;
var SUNDAY_HOURS_CELL = 9;
var LAST_SHIFT = ALL_SHIFTS.indexOf('1-2');
// ranges of all shifts in which tutoring is not offered.
var nonActiveShifts = ['B2:B25','B42:B49','G22:G49'];
var tutors = [];

// Spreadsheet properties
var spreadsheetId = '1E82wrm8FP9MftoKQkWQAK7ecz9LhoJdSSS2OiPzisGc';
var columns = ['B','C','D','E','F','G'];
var STARTING_ROW = 2;

//-------------------------- UI-related --------------------------
/**
 * Creates a menu entry in the Google Sheets UI when the document is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Schedule Helper')
    .addItem('Generate schedule', 'generateSchedule')
    .addItem('Check waitlist', 'showWaitlistSidebar')
    .addItem('Send emails', 'showEmailPrompt')
    .addToUi();
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 */
function showWaitlistSidebar() {
  var html = HtmlService.createTemplateFromFile('sidebar')
      .evaluate()
      .setTitle('Waitlist Sidebar');
  SpreadsheetApp.getUi().showSidebar(html);
}


/**
 * Opens a user prompt that accepts the name of a tutor to email.
 */
function showEmailPrompt() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
      'Please input the name of the tutor you want to email',
      ui.ButtonSet.OK_CANCEL);
  // Process the user's response.
  var button = result.getSelectedButton();
  var name = result.getResponseText().trim();
  if (button == ui.Button.OK) {
    sendEmailTo(name);
  }
}

/**
 * Used for inserting templated HTML in the sidebar.
 */
function include(file) {
  return HtmlService.createHtmlOutputFromFile(file).getContent();
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('sidebar');
}

//-------------------------- Data reading --------------------------
/**
 * Constructs a tutor object for each survey response and returns a list of 
 * all tutor objects in the order his/her response was submitted.
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

//--------------------------- Schedule automation and display --------------------------
function generateSchedule() {
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
  // display table of available and assigned hours
  listNumberOfHours('givenHours', tutors, SCHEDULE_SHEET, 'I', 'J');
  listNumberOfHours('assignedHours', tutors, SCHEDULE_SHEET, 'L', 'M');
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
  ALL_SHIFTS.forEach(function(shift) {
    sheet.getRange(leftColumn+shiftRow)
      .setValue(shift)
      .setFontWeight('bold')
      .setHorizontalAlignment('right');
    shiftRow += MAX_TUTORS;  
  });
  // per shift per day
  var shiftBlocks = getAllShiftRanges();
  shiftBlocks.forEach(function(block) {
    sheet.getRange(block.range)
      .setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  });
}

/** 
 * Writes a two-columned table for each tutor, displaying his/her name
 * and the number of [type = given or assigned] hours.
 */
function listNumberOfHours(type, tutors, scheduleName, nameCol, hourCol) {
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(scheduleName);
  var row = STARTING_ROW;
  var nameCell = sheet.getRange(nameCol+row);
  var hourCell = sheet.getRange(hourCol+row);
  tutors = sortByLastName_(tutors);
  for (var i = 0; i < tutors.length; i++) {
    nameCell.setValue(tutors[i].name);
    hourCell.setValue(tutors[i][type]);
    row++;
    nameCell = sheet.getRange(nameCol+row);
    hourCell = sheet.getRange(hourCol+row);
  }
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
 * ex.) schedule = [['Ash Ketchum', 'Philip Fry'], ...]
 * where schedule[0] represents Sunday tutors, schedule[1] represents Monday tutors, etc.
 */
function createSchedule(tutors, dayOfWeek) {
  var schedule = [];
  for (var i = 0; i < ALL_SHIFTS.length; i++) {
    schedule[i] = [];
    // no more active shifts after Friday, 1-2 pm 
    if (dayOfWeek === 'friday' && i > LAST_SHIFT) {
      break;
    }
    tutors.forEach(function(tutor) {
      if (tutor.shifts[dayOfWeek].indexOf(ALL_SHIFTS[i]) != -1) {
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
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(SCHEDULE_SHEET);
  var cell = sheet.getRange(column+row);
  for (var count = 0; count < MAX_TUTORS; count++) {
    if (!tutorsPerShift[count]) {
      cell.setValue('');
    } else {
      var tutor = tutorsPerShift[count];
      // write the tutor's name in the shift's cell
      cell.setValue(tutor.name);
      // increment the assigned hours of the given tutor
      if (tutor.assignedHours === undefined) {
        tutor.assignedHours = 1;
      } else {
        tutor.assignedHours++;
      }
    }
    row++;
    cell = sheet.getRange(column+row);
  }
}

//-------------------------- Waitlisting --------------------------
/**
 * Returns the names of the tutors that are waitlisted for the shift currently
 * being highlighted by the user.
 */
function findWaitlistForHighlight() {
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
 * Returns the names of the tutors that are waitlisted for the shift 
 * specified by the sidebar dropdowns.
 */
function findWaitlistForSelection(day, shift) {
  day = day.toLowerCase();
  if (true) {
    // return list of tutors that are waitlisted for the selected range.
    var waitlist = getWaitlistedTutors(range, day, shift);
    var waitlistNames = waitlist.map(function(tutor) { return tutor.name });
    return waitlistNames;
  } else {
    throw new Error('Invalid values ' + day + ', ' + shift + ' selected.');
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
  var currentTutors = getTutorsOnSchedule(range);
  var waitlist = tutors.filter(function(tutor) {
    var notOnShift = currentTutors.indexOf(tutor.name) === -1;
    var canWorkShift = tutor.shifts[day].indexOf(shift) >= 0;
    return (notOnShift && canWorkShift);
  });
  return waitlist;
}

function getTutorsOnSchedule(range) {
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

//-------------------------- Email notifications --------------------------
/**
 * Sends an email to the tutor that matches the given name with 
 * the shifts that he/she is scheduled to work.
 */
function sendEmailTo(name) {
  fetchSurveyData();
  var tutor = tutors.filter(function(tutor) {
    return tutor.name === name;
  })[0];
  var assignedShifts = getAssignedShifts(name, "Final Schedule");
  var body = constructBodyFromShiftData_(assignedShifts);
  GmailApp.sendEmail(tutor.email, "ESS Tutoring: Shifts for " + tutor.name, body); 
}

function constructBodyFromShiftData_(assignedShifts) {
  var result = 'Listed below are the shifts you are scheduled to work for the upcoming semester:\n';
  daysOfTheWeek.forEach(function(day) {
    if (assignedShifts[day] !== undefined) {
      result += day.charAt(0).toUpperCase() + day.slice(1) + ': '
             + arrayToString(assignedShifts[day])
             + '\n';
    }
  })
  return result;
}

/**
 * Returns an object with properties equal to the days of the week.
 * Each property is an array containing a string representing the shift time
 * that the tutor is assigned to work (aka where the tutors name is written on the given sheet name).
 * ex.) assignedShifts = {sunday: ['9-10', '10-11'], monday: ['3-4'], ...}
 */
function getAssignedShifts(tutorName, sheetName) {
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  // shiftRanges is an array of all the ranges representing a single shift in A1 notation.
  var shiftRanges = getAllShiftRanges();
  var assignedShifts = new Object();
  for (var i = 0; i < shiftRanges.length; i++) {
    // tutorsInRange is an array of the names of the tutors that exist in the given range. 
    var tutorsInRange = getTutorsOnSchedule(shiftRanges[i].range);
    // if the tutorName is in the given range, then include the shift's day and time in the results.
    if (tutorsInRange.indexOf(tutorName) !== -1) {
      var currentDay = shiftRanges[i].day;
      var currentShift = shiftRanges[i].time;
      if (assignedShifts[currentDay] === undefined) {
        assignedShifts[currentDay] = [];
      }
      // ex) assignedShifts.sunday = [9-10, 10-11, etc.]
      assignedShifts[currentDay].push(currentShift);
    }
  }
  return assignedShifts;
}

//-------------------------- General-purpose / Formatting --------------------------
function arrayToString(array) {
  var string = '';
  for (var i = 0; i < array.length; i++) {
    string += array[i];
    if (i !== array.length-1) string += ", ";
  }
  return string;
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
    for (var j = 0; j < ALL_SHIFTS.length; j++) {
      var shift = new Object();
      shift.range = column+startRow + ':' + column+endRow;
      shift.day = daysOfTheWeek[i];
      shift.time = ALL_SHIFTS[j];
      // store the range of the current shift block
      allShiftRanges.push(shift);
      startRow += MAX_TUTORS;
      endRow += MAX_TUTORS;
    }
  }
  return allShiftRanges;
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


function sortByGivenHours_(tutors) {
  tutors.sort(function(a,b) {
    return a.givenHours - b.givenHours;
  });
  return tutors; // bc apparently js doesn't pass by reference >:(
}

function sortByLastName_(tutors) {
  tutors.sort(function(a,b) {
    var aName = a.name.split(' ');
    var aLast = aName[aName.length-1];
    var bName = b.name.split(' ');
    var bLast = bName[bName.length-1];
    if (aLast < bLast) {
      return -1;
    } else if (aLast > bLast) {
      return 1;
    }
    return 0;
  });
  return tutors;
}