letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
weekdays = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"];
var sheet = SpreadsheetApp.getActive();

// Set User Menu
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
    .createMenu('Time Tracker')
    .addItem('Fill Category and Calendar Times', "fillCategoryAndCalendarTimes")
    .addItem('Fill Category Times', "fillCategoryTimes")
    .addItem('Fill Calendar Times', "fillCalendarTimes")
    .addItem('Fill Calendar Row Colors', "fillCalendarRowColors")
    .addToUi();
}

///////////////////////////////////////////////////////////////
///                   Helper Functions
///////////////////////////////////////////////////////////////

// returns the type of the string
var getType = function(obj) {
  return ({}).toString.call(obj).match(/\s([a-zA-Z]+)/)[1].toLowerCase()
}

// checks that the first and second string are equal
function equalsString(v, check){
  if ((getType(v) == 'string') && v.toLowerCase() == check ) return true;
  else return false;
}

// get the value at a sheet cell
function getSheetVal(cell){
  return sheet.getRange(cell).getValue();
}

// set the value at a sheet cell
function setSheetVal(cell, val){
  sheet.getRange(cell).setValue(val);
}

// creates a date array from given month, date, and year cells
function getDateFromCells(monthCell, dateCell, yearCell){
  var month = getSheetVal(monthCell);
  var day = getSheetVal(dateCell);
  var year = getSheetVal(yearCell);
  return [month, day, year];
}

// return true if a cell in the date array is empty
function emptyDateCell(date){
  for (var i = 0; i < date.length; ++i){
    if (date[i] == "") return true;
  }
  return false;
}

// updates the array with a string. returns false upon non string
function updateArrayWString(array, sheet, cell, iterator, end){
  val = getSheetVal(cell) + "";
  val = val.toLowerCase();
  if ( equalsString(val, "") ) return false;
  if ( equalsString(val, "total") ) return false;
  if (end && end == iterator) return false;
  array.push(val +"");
  return true;
}

// updates the array with a number. returns false upon non number
function updateArrayWNumber(array, sheet, cell, iterator, end){
  val = getSheetVal(cell);
  if ( val != 0 && !val ) return false;
  if (getType(val) != "number") return false;
  if (end && end == iterator) return false;
  array.push(val);
  return true;
}

/**
 * Starting at a certain cell, collects all values vertically until it reaches endRow, an empty cell, or a cell that has "total" as its value
 * @param  {string} column   
 * @param  {int} startRow 
 * @param  {int} endRow   
 * @return {array}          all values collected
 */
function getVerticalCellValues(column, startRow, updater, endRow){
  var vals = [];
  var val;
  var it = startRow;
  
  while (true){
    var cell = column + it;
    
    if (!updater) updater = updateArrayWString;
    if (updater(vals, sheet, cell, it, endRow)) ++it;
    else break;

  }
  return vals;
}

// this is effectively the same as getVerticalCellValues, except that it collects values across a row starting at startCol and ending at endCol
function getHorizantalCellValues(row, startCol, updater, endCol){
  var vals = [];
  var val;
  var it = letters.indexOf(startCol);
  
  while (true){
    var letter = letters[it];
    var cell = letter + row;
    
    if (!updater) updater = updateArrayWString;
    if (updater(vals, sheet, cell, it, endCol)) ++it;
    else break;
  }
  return vals;
}

/**
 * returns date object from array
 * @param  {array} array [month, date, year]
 * @return {date}       date object
 */
function dateFromArray(array){
  var month = array[0];
  var day = array[1];
  var year = array[2];

  return new Date(array[2], array[0] - 1, array[1], 0, 0, 0);

}

///////////////////////////////////////////////////////////////

/**
 * return the start date. either the value in default cells or current time if error
 * @return {date}
 */
function getStartDate(){
  var date = getDateFromCells("F7", "G7", "H7");
  if (emptyDateCell(date)) return new Date();
  else return dateFromArray(date);
}

/**
 * helper function for function below
 * @param  {array} date array with date values
 * @return {date}      next date
 */
function defaultEnd(date){

  var next;
  if (getType(date[0]) == "string"){
    next = date[0].toLowerCase();
  }
  else next = "saturday";
  var indexOfNext = weekdays.indexOf(next);

  var date = new Date();

  var difference = indexOfNext - date.getDay();
  if (difference <= 0) difference += 7;

  date.setDate(date.getDate() + difference);

  return date;
}

/**
 * returns the end date
 * if date with correct input given, that date is used
 * can give it as a date array [month, date, year]
 * or as [day_of_the_week]
 * e.g. 12 | 28 | 2015 will set the date to 12/28/2015
 *      Monday will set the date as next monday
 * if none of the conditions above are satisfied, then the date is set as next saturday
 * @return {date} end date
 */
function getEndDate(){
  var date = getDateFromCells("F8", "G8", "H8");
  if (emptyDateCell(date)) return defaultEnd(date);
  else return dateFromArray(date);
}

/**
 * starting at B2, going horizantally, return a list of sub categories 
 * @return {array of strings} 
 */
function getSubCategoryList(){
  return getHorizantalCellValues(2, 'B');
}

/**
 * starting at B8, going vertically, return list of subCalendars
 * @return {array of strings}
 */
function getSubCalendarList(){
  return getVerticalCellValues('B', 8);
}

/**
 * starting at A8, going vertically, return list of colors for subCalendars. also used as a proxy for which subCalendars will be used to track time delegation for subCategories
 * @return {array of strings}
 */
function getSubCalendarColors(){
  return getVerticalCellValues('A', 8);
}

/**
 * fills in map that associates subCategories with time assignment. the key of the map is the subCategory and the value is the number of hours assigned to that subCategory
 * @param  {google calendar} calendar         
 * @param  {map} subCategoryTimes 
 * @param  {array of strings} subCategories    
 * @param  {Date} startDate        
 * @param  {Date} endDate          
 */
function getSubCategoryTimes(calendar, subCategoryTimes, subCategories, startDate, endDate){

  var events = calendar.getEvents(startDate, endDate);
  
  for (var i = 0; i < events.length; i++){
    
    var ev = events[i];
    var eventTitle = ev.getTitle().toLowerCase();
    
    for (var j = 0; j < subCategories.length; ++j){
      
      var subCategory = subCategories[j];
      if (eventTitle.indexOf(subCategory) < 0) { continue; }
      
      var time = ev.getEndTime() - ev.getStartTime();
      time /= (60*60*1000);
      
      if ( subCategoryTimes[subCategory] != 0 && !subCategoryTimes[subCategory] ) 
        subCategoryTimes[subCategory] = 0;
      subCategoryTimes[subCategory] += time;
    }
  }
}

/**
 * fills in the row at B3 with the time assigned to different subCategories
 */
function fillCategoryTimes(){
  
  var startDate = getStartDate();
  var endDate = getEndDate();

  var calendars = CalendarApp.getAllCalendars();
  var nColors = getSubCalendarColors().length;
  var subCalendars = getSubCalendarList().slice(0, nColors);

  var subCategories = getSubCategoryList();
  var subCategoryTimes = {};

  
  for (var i = 0; i < calendars.length; ++i){
    var calendar = calendars[i];

    var name = calendar.getName().toLowerCase();
    if (subCalendars.indexOf(name) == -1){ 
      continue;
    }

    getSubCategoryTimes(calendar, subCategoryTimes, subCategories, startDate, endDate)
  
  }

  var startingLetterIndex = letters.indexOf("B");
  
  Logger.log(subCategories);
  for (var i = 0; i < subCategories.length; ++i){
  
    var letter = letters[i+startingLetterIndex]
    var cell = letter + 3;

    var subCategory = subCategories[i];
    var subCategoryTime = subCategoryTimes[subCategory];
    if (!subCategoryTime)
      subCategoryTime = 0;
    else subCategoryTime = Math.round(subCategoryTime * 100) / 100;

    Logger.log(subCategory + " : " + subCategoryTime);
    setSheetVal(cell, subCategoryTime);
    sheet.getRange(cell).setBackground("yellow")
  }

}

/**
 * finds the time in assigned to a particular subCalendar
 * @param  {google calendar} calendar         
 * @param  {Date} startDate
 * @param  {Date} endDate
 */
function timeInSubCalendar(calendar, startDate, endDate){
  
  var events = calendar.getEvents(startDate, endDate);
  var sum = 0;
  for (var i = 0; i < events.length; i++){
    var ev = events[i];

    var time = ev.getEndTime() - ev.getStartTime();
    time /= (60*60*1000);
    sum += time;
  }
  return sum;
}

/**
 * finds the time assigned to each subCalendar
 * @param  {google calendar} calendar         
 * @param  {Date} startDate
 * @param  {Date} endDate
 */
function getSubCalendarTimes(subCalendarNames, startDate, endDate){
  var subCalendarObjs = [];
  
  var subCalendars = CalendarApp.getAllCalendars();

  for (var i = 0; i < subCalendars.length; i++){
    var subCalendar = subCalendars[i];
    var calName = subCalendar.getName().toLowerCase();

    var indx = subCalendarNames.indexOf(calName);
    if (indx == -1){ 
      continue;
    }

    var row = indx + 8;
    var time = timeInSubCalendar(subCalendar, startDate, endDate);
    
    var subCalendarObj = {
      Time: time,
      Row: row
    }
    
    subCalendarObjs.push(subCalendarObj);
  }

  return subCalendarObjs;
}

/**
 * fills in the column starting at C8 with the number of hours delegated to different subCalendars
 */
function fillCalendarTimes(){

  var startDate = getStartDate();
  var endDate = getEndDate();
  
  var subCalendarNames = getSubCalendarList();

  var subCalendarObjs = getSubCalendarTimes(subCalendarNames, startDate, endDate);

  var startingRow = 8;
  for (var i = 0; i < subCalendarObjs.length; ++i){
    var time = subCalendarObjs[i].Time;
    time = Math.round(time * 100) / 100;

    var cell = "C" + subCalendarObjs[i].Row;
    setSheetVal(cell, time);
  }
}

/**
 * based on the colors given in the column at A8, A:C from 8 onwards is colored
 */
function fillCalendarRowColors(){
  var colors = getSubCalendarColors();

  var colStart = "A";
  var colEnd = "C";
  var startRow = 8;
  for (var row = 0; row < colors.length; ++row){
    var col1 = colStart + (startRow + row);
    var col2 = colEnd + (startRow + row);
    var range = col1+":"+col2;

    sheet.getRange(range).setBackground(colors[row])
  }
}

/**
 * wrapper function
 */
function fillCategoryAndCalendarTimes(){
  fillCategoryTimes()
  fillCalendarTimes()
}