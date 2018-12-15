//-------------------------- CONSTANTS --------------------------
// Names
var SURVEY_SHEET = 'Form Responses 1';
var SCHEDULE_SHEET = 'New Schedule';
var INDIVIDUAL_SCHEDULE = 'Your Tutoring Schedule';

// Schedule properties
var DAYS_OF_THE_WEEK = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday'];
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
var columns = ['B','C','D','E','F','G'];
var STARTING_ROW = 2;

//-------------------------- UI-related --------------------------
/**
 * Triggered when the user first installs the add-on; populates add-on menu.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Creates a menu entry in the Google Sheets UI when the document is opened.
 */
function onOpen(e) {
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
      'Please input the name(s) of the tutor you want to email',
      ui.ButtonSet.OK_CANCEL);
  // Process the user's response.
  var button = result.getSelectedButton();
  var names = result.getResponseText().split(',');
  var tutorsToEmail = [];
  if (button == ui.Button.OK) {
    fetchSurveyData();  // fetch data from responses once.
    // first, run checks to ensure all names are valid.
    names.forEach(function(name) {
      // find tutor associated with current name.
      var tutor = tutors.filter(function(tutor) {
        return tutor.name === name;
      })[0];
      // if a given name does not match a tutor's name, then the input is invalid.
      if (!tutor) {
        ui.alert(name + ' is not a valid name.');
        return;
      }
      tutorsToEmail.push(tutor);
    });
    // if the input was completely valid, send emails to all tutors specified by the user.
    tutorsToEmail.forEach(function(tutor) {
      sendEmailTo(tutor);
    });
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
  const survey = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SURVEY_SHEET);
  const lastRow = survey.getLastRow();
  const infoRange = survey.getRange('A2:F' + lastRow);
  const basicInfo = infoRange.getValues();
  const timeRange = survey.getRange('I2:R' + lastRow);
  const hoursInfo = timeRange.getValues();
  if (!basicInfo) {
    Logger.log('No data found.');
  } else {
    for (var row = 0; row < basicInfo.length; row++) {
      var tutor = new Object();
      tutor.row = row;
      tutor.timestamp = basicInfo[row][0]; // A
      if (tutor.timestamp === '') { // skip all empty rows.
        continue;
      }
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
      // assigned hours are 0 before the schedule is written.
      tutor.assignedHours = 0;
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
  for (var day = 0; day < DAYS_OF_THE_WEEK.length; day++) {
    var shiftsPerDay = tutor.shifts[DAYS_OF_THE_WEEK[day]];
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
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  try {
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName(SCHEDULE_SHEET));
  } catch(e) {
    // insert after form responses
    spreadsheet.insertSheet(SCHEDULE_SHEET, 1);
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName(SCHEDULE_SHEET));
  }
  writeBlankSchedule(SCHEDULE_SHEET);
  fetchSurveyData();
  tutors = sortByGivenHours_(tutors);
  // construct schedule
  for (var i = 0; i < DAYS_OF_THE_WEEK.length; i++) {
    var currentDayTutors = createSchedule(tutors, DAYS_OF_THE_WEEK[i]);
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
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
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
      .setNumberFormat('@STRING@') // to avoid reading shift hours as dates
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
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(scheduleName);
  // set title at first row.
  sheet.getRange(nameCol+1).setValue(type).setFontWeight('bold');
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
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SCHEDULE_SHEET);
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
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SCHEDULE_SHEET);
  var cell = sheet.getRange(column+row);
  for (var count = 0; count < MAX_TUTORS; count++) {
    if (!tutorsPerShift[count]) {
      cell.setValue('');
    } else {
      var tutor = tutorsPerShift[count];
      // write the tutor's name in the shift's cell
      cell.setValue(tutor.name);
      // increment the assigned hours of the given tutor
      tutor.assignedHours++;
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
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SCHEDULE_SHEET);
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
function findWaitlistForSelection(day, time) {
  var allShiftBlocks = getAllShiftRanges();
  // find the range of the shift that matches the given day and time.
  var range = allShiftBlocks.filter(function(shiftBlock) {
    return shiftBlock.day === day && shiftBlock.time === time;
  })[0].range;
  if (true) {
    // return list of tutors that are waitlisted for the selected day and shift.
    var waitlist = getWaitlistedTutors(range, day, time);
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
    if (range === allValidRanges[i].range) return true;
    // TODO: add check if in non active shift
  }
  return false;
}

function getDayFromRange(range) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SCHEDULE_SHEET);
  var column = sheet.getRange(range).getColumn();
  return sheet.getRange(1, column).getValue().toLowerCase();
}

function getShiftFromRange(range) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SCHEDULE_SHEET);
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
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SCHEDULE_SHEET);
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
 * Sends an email to the given tutor with the shifts that he/she is scheduled to work.
 */
function sendEmailTo(tutor) {
  // TODO: allow user to specify sheet name.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Final Schedule");
  var assignedShifts = getAssignedShifts(name, sheet);
  var schedule = produceScheduleFor(name, sheet);
  var body = constructBodyFromShiftData_(assignedShifts, schedule);
  GmailApp.sendEmail(tutor.email, "ESS Tutoring: Shifts for " + name, body); 
}

function constructBodyFromShiftData_(assignedShifts, schedule) {
  var result = 'Listed below are the shifts you are scheduled to work for the upcoming semester:\n';
  DAYS_OF_THE_WEEK.forEach(function(day) {
    if (assignedShifts[day] !== undefined) {
      result += day.charAt(0).toUpperCase() + day.slice(1) + ': '
             + arrayToString(assignedShifts[day])
             + '\n';
    }
  });
  result += "\n\nView your schedule here: " + getSheetUrl(schedule);
  return result;
}

/**
 * Returns an object with properties equal to the days of the week.
 * Each property is an array containing a string representing the shift time
 * that the tutor is assigned to work (aka where the tutors name is written on the given sheet name).
 * ex.) assignedShifts = {sunday: ['9-10', '10-11'], monday: ['3-4'], ...}
 */
function getAssignedShifts(tutorName, sheet) {
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

/**
 * Creates a sheet in a separate spreadsheet that is a copy of the given schedule
 * with only tutorName highlighted.
 * 
 * @param {string} the name of the tutor the individual schedule is intended for.
 * @param {string} the name of schedule we want to base the individual copy on.
 * @return {Sheet} the newly created sheet containing tutorName's schedule.
 */
function produceScheduleFor(tutorName, schedule) {
  var ss = addIndividualSpreadsheet(tutorName);
  var copy = schedule.copyTo(ss);
  // rename the duplicated schedule sheet (must delete first if already exists).
  if (ss.getSheetByName(INDIVIDUAL_SCHEDULE) !== undefined) {
    ss.deleteSheet(ss.getSheetByName(INDIVIDUAL_SCHEDULE));
  }
  copy.setName(INDIVIDUAL_SCHEDULE);
  // delete all other existing sheets.
  ss.getSheets().forEach(function(sheet) {
    if (sheet.getName() !== INDIVIDUAL_SCHEDULE) {
      ss.deleteSheet(sheet);
    }
  });
  highlightSchedule(tutorName, copy);
  // remove all non-schedule data (such as list of assigned hours)
  clearNonScheduleData(copy);
  return copy;
}

/**
 * Given a tutor's name, creates a new sheet in the "Individual Schedules"
 * Spreadsheet titled "(tutorName) Tutoring Schedule."
 * Allows anyone with a link to view the newly created file.
 * 
 * @return {Sheet} the newly created Spreadsheet.
 */
function addIndividualSpreadsheet(tutorName) {
  var folder = DriveApp.getFoldersByName("Individual Schedules").next();
  // check if the spreadsheet for this tutor already exists
  var ssName = "(" + tutorName + ") Tutoring Schedule";
  var iterator = folder.getFilesByName(ssName);
  while (iterator.hasNext()) { // already exists
    var oldSS = SpreadsheetApp.open(iterator.next());
    return oldSS;
  }
  // does not already exist, so create a new spreadsheet.
  var newSS = SpreadsheetApp.create(ssName);
  // work-around for moving a new file to a specific folder.
  var temp = DriveApp.getFileById(newSS.getId());
  folder.addFile(temp);
  DriveApp.getRootFolder().removeFile(temp);
  return newSS;
}

/**
 * Given a tutor's name and a Sheet representing a schedule, modifies the given
 * Sheet so all occurances of the tutor's name is highlighted.
 */
function highlightSchedule(tutorName, sheet) {
  // a more complicated strategy, but less expensive in terms of API calls.
  var shiftBlocks = getAllShiftRanges();
  for (var i = 0; i < shiftBlocks.length; i++) {
    var currentRange = shiftBlocks[i].range;
    // getValues() returns a 2D array, where values[0] = [tutors in row]
    var tutorsInShift = sheet.getRange(currentRange).getValues();
    for (var j = 0; j < MAX_TUTORS; j++) {
      // only check first element because there will only be one value per array (range is incremented by row)
      if (tutorsInShift[j][0] === tutorName) {
        // calcuate cell position based on loop iteration.
        var col = currentRange.charAt(0);
        var rangeString = JSON.stringify(currentRange);
        var firstRowInRange = parseInt(currentRange.substring(1, currentRange.indexOf(':')));
        var row = firstRowInRange + j;
        var cell = sheet.getRange(col+row);
        // mark the cell by highlighting.
        cell.setBackground('#ffff7f').setFontWeight('bold');
      }
    }
  }
  // simpler strategy, but makes one call to getValue() per cell in the schedule.
  /*
  for (var row = STARTING_ROW; row < sheet.getLastRow(); row++) {
    columns.forEach(function(column) {
      var cell = sheet.getRange(column+row);
      if (cell.getValue() === tutorName) {
        cell.setBackground('#ffff7f').setFontWeight('bold');
      }
    });
  }
  */
}

/**
 * Clears all ranges that contain content outside of the written schedule.
 */
function clearNonScheduleData(sheet) {
  var nonScheduleCol = String.fromCharCode(columns[columns.length-1].charCodeAt() + 1);
  var nonScheduleRow = (ALL_SHIFTS.length+1) * MAX_TUTORS + STARTING_ROW;
  var lastCol = String.fromCharCode(sheet.getLastColumn() + 'A'.charCodeAt());
  // vertical coverage (right of the schedule)
  sheet.getRange(nonScheduleCol + 1 + ":" + lastCol + sheet.getLastRow()).clearContent();
  // horizontal coverage (below the schedule)
  sheet.getRange('A' + nonScheduleRow + ":" + nonScheduleCol + sheet.getLastRow()).clearContent();
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
 * Returns an array containing objects that represent a single shift block;
 * Each shift block has the following properties: 
 * - range: ranges in A1 format ('B2:B5')
 * - day: day of the given shift block ('sunday')
 * - time: hours of the given shift block ('9-10')
 */
function getAllShiftRanges() {
  var allShiftRanges = [];
  for (var i = 0; i < DAYS_OF_THE_WEEK.length; i++) {
    var column = columns[i];
    var startRow = STARTING_ROW;
    var endRow = startRow + MAX_TUTORS-1; // subtract one because the group of cells is inclusive
    for (var j = 0; j < ALL_SHIFTS.length; j++) {
      var shift = new Object();
      shift.range = column+startRow + ':' + column+endRow;
      shift.day = DAYS_OF_THE_WEEK[i];
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
      shifts[DAYS_OF_THE_WEEK[day]] = [];
    } else {
      shifts[DAYS_OF_THE_WEEK[day]] = formatHours_(hoursPerDay[day]);
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

function getSheetUrl(sheet) {
  return "https://drive.google.com/file/d/"+ sheet.getParent().getId() +"/view";
}