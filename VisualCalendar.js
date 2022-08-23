/*****
 * 
 * The functions below allow for the updating of the visual based on the bookings sheet.
 * It works both "onFormSubmit", but can also be run against the whole booking sheet
 * 
 ****/
 //TODO add cellStatusText for out of service and out of order, so you can add a booking with those as status and it blocks the calendar
 //TODO add booking status of Interested to keep track of all those interested and in what dates

function updateVisualFromBookingSheet() {

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var bookingSheet = spreadsheet.getSheetByName("Bookings");
  bookingDataRange = bookingSheet.getDataRange();
  bookingValues = bookingDataRange.getValues();

  visualSheet = spreadsheet.getSheetByName("new gen-mar 21 ");
  visualDataRange = visualSheet.getDataRange();
  visualDataValues = visualDataRange.getValues();

  //Check how many months are in the sheet
  if (timing) console.time("getMonthsInSheet");
  var numOfMonths = getMonthsInSheet(visualDataValues);
  if (timing) console.timeEnd("getMonthsInSheet");

  if (timing) console.time("clearVis");
  clearOldVisuals(numOfMonths);
  if (timing) console.timeEnd("clearVis");
  //Utilities.sleep(5000);

  var event = {};

  if (timing) console.time("processingAllBookings");
  //log("bookingValues len " + bookingValues.length + " -- bookingValues[0].len " + bookingValues[0].length);
  for (var i = 1; i < bookingValues.length; i++) {
    log("Processing event: " + i);
    for (var j = 0; j < bookingValues[i].length; j++) {
      if (bookingValues[0][j] === "") break;
      event[bookingValues[0][j]] = bookingValues[i][j].valueOf();
    }
    if (timing) console.time("form submit");
    onFormSubmit(event, i);
    if (timing) console.timeEnd("form submit");
    event = {};
  }
  if (timing) console.timeEnd("processingAllBookings");
}

function clearOldVisuals(numOfMonths) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("new gen-mar 21 ");

  var startRow = 4;

  if (numOfMonths == -1) {
    throw new Error("Didn't find months in sheet, formatting issue.");
  }

  //log("numOfMonths " + numOfMonths + " -- rowsBetweenMonths " + rowsBetweenMonths);

  if (timing) console.time("getMonthsInSheet-clear");
  for (var i = 0; i < numOfMonths; i++) {
    //log("range " + (startRow + (rowsBetweenMonths * i)))
    var range = sheet.getRange((startRow + (rowsBetweenMonths * i)), 2, 11, 31);
    //log("range a1notation: " + range.getA1Notation());
    range.clearContent();
    range.clearFormat();
    range.setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
  if (timing) console.timeEnd("getMonthsInSheet-clear");
}

function getMonthsInSheet(dataValues) {
  var month = -1;

  for (var i = 0; i < dataValues.length; i++) {
    var cellValue = dataValues[i][0];

    if (!isNaN(cellValue) && cellValue != "") {
      month = cellValue;
    }

    if (cellValue == 2) {
      rowsBetweenMonths = i;
    }

    //log("cellValue " + cellValue + "month " + month + " -- dataValues.length " + dataValues.length);
  }
  //log("rowsBetweenMonths " + rowsBetweenMonths + " --typeof " + typeof rowsBetweenMonths);
  return month;
}

function onFormSubmit(event, bookingRow) {

  //sample event namedValues
  /*
  var event = {
      DataArrivo: "8/19/2021", 
      Metodo: "Bonifico", 
      DataPartenza: "8/23/2021", 
      Acconto: "0", 
      Note: "mai pagato", 
      //CestoBio: "si", 
      Timestamp: "8/9/2021 15:48:06", 
      //Ebike:"si", 
      Email: "loser@dontpay.com", 
      //Voucher: "", 
      //Massagi: "2", 
      //Aperitivo: "Simplice", 
      BookingStatus: "Cancellato", 
      Telefono: "000 00 0000", 
      Camera: "Rose", 
      Contatti: "Loser Dont Pay", 
      Dessert: "si"
    }
    
  */
  log(event);
  log(event.namedValues);

  var visualSheetId = "#gid=1603092828";
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //TODO update name to more appropriate name
  //TODO handle year
  var visualSheet = spreadsheet.getSheetByName("new gen-mar 21 ");
  var bookingSheet = spreadsheet.getSheetByName("Bookings");


  var eventValues = null;
  if (event.namedValues) {
    eventValues = event.namedValues;
  } else {
    eventValues = event;
  }
  
  //if (timing) console.time("getVisualDataRange");
  
  //checking if variables exist to understand if we are coming from onFormSubmit or from updateVisualsFromSheet
  if (!visualDataRange) visualDataRange = visualSheet.getDataRange();
  if (!visualDataValues) visualDataValues = visualDataRange.getValues();
  if (!bookingDataRange) bookingDataRange = bookingSheet.getDataRange();
  if (!bookingValues) bookingValues = bookingDataRange.getValues();
  if (rowsBetweenMonths == 0) getMonthsInSheet(visualDataValues);
  
  //if (timing) console.timeEnd("getVisualDataRange");

  var cellStatusText = getStatusVisualText(eventValues["BookingStatus"].toString().toLowerCase());
  var visualUpdateObject = getVisualUpdateObject(cellStatusText, bookingRow, rowsBetweenMonths, eventValues, visualDataValues, bookingSheet, visualSheetId);

  bookingDataRange.getCell(bookingRow, 1).setValue(new Date());

  if (timing) console.time("updatedCells");
  //Update the cells to show the booking, with the hyperlink and appropriate borders

  //If the visual short text is not empty (confermato, da pagare, closed, out of order or out of service), then ignore 
  if(cellStatusText != "") {
    updateVisualForBooking(visualUpdateObject, visualDataRange);
  }

  if (timing) console.timeEnd("updatedCells");
}

function getVisualUpdateObject(cellStatusText, bookingRow, rowsBetweenMonths, eventValues, visualDataValues, bookingSheet, visualSheetId) {
  var visualObject = {};

  var monthRow = 0;
  var roomRow = 0;

  visualObject.dateOfArrival = new Date(eventValues["DataArrivo"]);
  visualObject.roomBooked = eventValues["Camera"].toString();
  visualObject.dateOfDeparture = new Date(eventValues["DataPartenza"]);
  visualObject.bookingStatus = eventValues["BookingStatus"].toString().toLowerCase();
  log("visualObject.bookingStatus " + visualObject.bookingStatus);

  if (!(visualObject.bookingStatus === "confermato" || visualObject.bookingStatus === "da pagare")) {
    log("Booking status [" + visualObject.bookingStatus + "] is not paid or to be paid, so not adding to calendar.");
    return;
  }

  //if coming from a form, then get the bookingRow from the sheet's last row
  if (!bookingRow) {
    visualObject.bookingRow = bookingSheet.getDataRange().getLastRow() - 1;
    log("bookingRow: " + visualObject.bookingRow);
    //if coming from updateVisuals, use the provided bookingRow
  } else {
    visualObject.bookingRow = bookingRow;
    log("bookingRow: " + visualObject.bookingRow);
  }

  //Rounding cause seen an instance with nights = 4.925
  visualObject.nights = Math.round((visualObject.dateOfDeparture - visualObject.dateOfArrival) / (86400 * 1000));
  //log("nights " + nights);

  visualObject.arrivalMonth = visualObject.dateOfArrival.getMonth() + 1;
  visualObject.arrivalMonthLastDay = new Date(visualObject.dateOfArrival.getFullYear(), visualObject.dateOfArrival.getMonth() + 1, 0).getDate();
  //log("arrivalMonthLastDay " + visualObject.arrivalMonthLastDay);
  visualObject.arrivalDate = visualObject.dateOfArrival.getDate();
  log("arrivalMonth " + visualObject.arrivalMonth + " -- arrivalDate " + visualObject.arrivalDate);
  
  //TODO arrivalYear aka choose a different sheet

  //if (timing) console.time("getMonthRow-DataRange");
  visualObject.monthRow = getIndexFromDataRange(visualDataValues, 0, -1, visualObject.arrivalMonth) + 1; //for index 0
  //log("monthRow " + visualObject.monthRow);
  //if (timing) console.timeEnd("getMonthRow-DataRange");
  
  //log("getting Index from col 0 till row " + visualDataValues.length + " searching for arrivalMonth " + arrivalMonth + " -- result: " + monthRow);
  
  //if (timing) console.time("getDateColumn-DataRange");
  visualObject.dateColumn = getIndexFromDataRange(visualDataValues, -1, 1, visualObject.arrivalDate) + 1; //for index 0
  //log("dateColumn " + visualObject.dateColumn);
  //if (timing) console.timeEnd("getDateColumn-DataRange");

  //if (timing) console.time("getRoomRow");
  var roomRows = getIndexesFromRange(visualDataValues, 0, visualObject.roomBooked);

  log("roomBooked " + visualObject.roomBooked + " -- roomRows: " + roomRows.toString() + " -- rowsBetweenMonths " + rowsBetweenMonths );
  for (var j = 0; j < roomRows.length; j++) {
    //log("Math.floor(roomRows[j] / rowsBetweenMonths) + 1: " + (Math.floor(roomRows[j] / rowsBetweenMonths) + 1));
    if ((Math.floor(roomRows[j] / rowsBetweenMonths) + 1) == visualObject.arrivalMonth) {
      roomRow = roomRows[j];
      break;
    }
  }
  //if (timing) console.timeEnd("getRoomRow");
  visualObject.roomRow = roomRow;
  log("monthRow " + visualObject.monthRow + " -- dateColumn " + visualObject.dateColumn + " -- roomRow " + visualObject.roomRow + "[" + visualObject.roomBooked + "]");

  if (visualObject.roomRow == 0) throw new Error("visualSheet format incorrect, can't find correct room row.");

  //Get the appropriate cell text to output based on bookingStatus
  visualObject.cellStatusText = cellStatusText;
  //log("bookingStatus " + bookingStatus + " -- cellStatusText " + cellStatusText);

  visualObject.visualSheetId = visualSheetId;

  log("visualObject " + JSON.stringify(visualObject));
  return visualObject;
}

function updateVisualForBooking(visualUpdateObject, visualDataRange) {
  var nights = visualUpdateObject.nights;
  var dateColumn = visualUpdateObject.dateColumn;
  var roomRow = visualUpdateObject.roomRow;
  var cellStatusText = visualUpdateObject.cellStatusText;
  var bookingRow = visualUpdateObject.bookingRow;
  var arrivalMonthLastDay = visualUpdateObject.arrivalMonthLastDay;
  var visualSheetId = visualUpdateObject.visualSheetId;
  
  var daysInFirstMonth = 0;
  
  for (var z = 0; z < nights; z++) {

    if ((dateColumn - 1) + z <= arrivalMonthLastDay) {
      log("in update: z " + z + " -- dateColumn + z " + (dateColumn + z) + " -- nights " + nights);
      
      var cellForBooking = visualDataRange.getCell(roomRow, dateColumn + z);
      var cellVal = cellForBooking.getValue();//visualDataValues[roomRow][dateColumn + z];
      //log("cellForBooking.getA1Notation() " + cellForBooking.getA1Notation() + " -- cellVal [" + cellVal + "] cellForBooking.getValue() " + cellForBooking.getValue());
      
      //before writing anything, check that the cell value are empty otherwise mark as conflict
      var conflictCheckedStatusText = checkForConflict(cellVal, cellStatusText, bookingRow);
      if (nights > 1 && z > 0) cellVal = conflictCheckedStatusText; else cellVal = (createHyperlink(visualSheetId, bookingRow, conflictCheckedStatusText));
      cellForBooking.setValue(cellVal);

      //log("cellForBooking.getA1Notation() " + cellForBooking.getA1Notation() + " -- cellVal [" + cellVal + "] cellForBooking.getValue() " + cellForBooking.getValue());
      daysInFirstMonth = z + 1;

    } else {
      var cellForBooking = visualDataRange.getCell(roomRow + rowsBetweenMonths, z - daysInFirstMonth + 1 + 1); //+1 cause col index starts from 1, and +1 to skip the room label column
      var cellVal = cellForBooking.getValue();//visualDataValues[roomRow + rowsBetweenMonths][z - daysInFirstMonth + 1];
      //log("cellForBooking.getA1Notation() " + cellForBooking.getA1Notation() + " -- cellVal [" + cellVal + "] cellForBooking.getValue() " + cellForBooking.getValue());

      var conflictCheckedStatusText = checkForConflict(cellVal, cellStatusText, bookingRow)
      if (nights > 1) cellVal = conflictCheckedStatusText; else cellVal = (createHyperlink(visualSheetId, bookingRow, conflictCheckedStatusText));
      cellForBooking.setValue(cellVal);
      //log("cellForBooking.getA1Notation() " + cellForBooking.getA1Notation() + " -- cellVal [" + cellVal + "] cellForBooking.getValue() " + cellForBooking.getValue());
    }

    if (nights > 1) {
      log("z " + z + " -- nights " + nights + " -- daysInFirstMonth " + daysInFirstMonth + " -- dateColumn " + dateColumn + " -- arrivalMonthLastDay " + arrivalMonthLastDay + " -- dateColum " + dateColumn);
      //setBorder(top, left, bottom, right, vertical, horizontal) 
      //First cell, border top left and bottom, not right
      if (z == 0) cellForBooking.setBorder(true, true, true, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      //Last cell, border top bottom and right
      if (z == nights - 1) cellForBooking.setBorder(true, false, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      //If middle cell, only top and bottom
      if (z != 0 && z != nights - 1) cellForBooking.setBorder(true, false, true, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      //if last cell of first month, straddling row, then top, and bottom no right
      if ((dateColumn - 1) + z == arrivalMonthLastDay && daysInFirstMonth != z && z != nights -1) cellForBooking.setBorder(true, null, true, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      //if first cell in new month, straddling row, then top, bottom, no left, and right handled by above ifs
      if ((dateColumn - 1) + z == arrivalMonthLastDay + 1) cellForBooking.setBorder(true, false, true, null, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }

    //log("z is " + z + " --nights " + nights + " -- z == 0 " + (z == 0) + " -- z == nights - 1 " + (z == nights - 1) + " -- z != 0 && z != nights - 1 && daysInFirstMonth == z" + (z != 0 && z != nights - 1 && daysInFirstMonth == z) + " -- (dateColumn - 1) + z == arrivalMonthLastDay && daysInFirstMonth != z" + ((dateColumn - 1) + z == arrivalMonthLastDay && daysInFirstMonth != z) + " -- daysInFirstMonth: " + daysInFirstMonth);
  }
}

//Return the correct shortcut text for the visual sheet based on booking status
function getStatusVisualText(bookingStatus) {
  var cellStatusText = "";
  switch (bookingStatus) {
    case "confermato":
      cellStatusText = "x";
      break;
    case "da pagare":
      cellStatusText = "d";
      break;
    case "out of service":
      cellStatusText = "os";
      break;
    case "out of order":
      cellStatusText = "oo";
      break;
    case "closed":
      cellStatusText = "c";
      break;
    default:
      cellStatusText = "";
      break;
  }

  return cellStatusText;
}

function createHyperlink(visualSheetId, bookingRow, cellStatusText) {
  return hyperlink = "=HYPERLINK(\"" + visualSheetId + "range=A" + (bookingRow) + "\", \"" + cellStatusText + "\")";
}

//Checks the visual sheet for any existing values and updates the text and the booking status if there is a conflict
function checkForConflict(cellVal, cellStatusText, bookingRow) {

      //log("cellStatusText at start of conflict " + cellStatusText);
      if (cellVal !== "") {
        cellStatusText = cellStatusText + "?";

        //log("cellStatusText in conflict " + cellStatusText);
        var statusToUpdate = bookingDataRange.getCell(bookingRow + 1, 10);
        //log("statusToUpdate val " + statusToUpdate.getValue() + " -- a1: " + statusToUpdate.getA1Notation() );
        statusToUpdate.setValue("Review - Conflict");
      }

      return cellStatusText;
}

//Function to find a specific value within a range of values
function getIndexFromDataRange(rangeValues, column, row, searchCriteria) {
  //log("col " + column + " row " + row + " search " + searchCriteria);
  if (timing) var startTime = new Date().getMilliseconds();

  //If row provided, search through rows for the desired value
  if (row >= 0) {
    for (var i = 0; i < rangeValues[row].length; i++) {
      //log("COLUMNS - checking rangeValues[0]["+i+"] =[" + rangeValues[row][i] + "] equal " + searchCriteria);
      if (rangeValues[row][i] == searchCriteria) {
        //if (timing) log("colsLoop " + (new Date().getMilliseconds() - startTime));
        return i;
      }
    }
  }

  //If column provided, search through columns for the desired value
  if (column >= 0) {
    for (var i = 0; i < rangeValues.length; i++) {
      //log("ROWS - checking rangeValues["+i+"["+column+"]] =[" + rangeValues[i][column] + "] equal " + searchCriteria);

      if (rangeValues[i][column] == searchCriteria) {
        //if (timing) log("rowsLoop " + (new Date().getMilliseconds() - startTime));
        return i;
      }
    }
  }

  //throw error if can't find search criteria
  throw new Error("visualSheet format incorrect, can't find " + searchCriteria + " in range [" + rangeText + "].");
}

//get indexes of all the rooms 
//TODO can probably get rid of a loop here by doing the following
  //get rowsInBetweenMonths
  //get monthRow
  //only loop over monthRow + rowsInBetweenMonths, so we do it once and return the row of the correct room row directly
  //also eliminates the extra loop after getIndexesFromRange to find the correct roomRow
  function getIndexesFromRange(rangeValues, column, searchCriteria) {

  var results = [];

  //log("rangeValues : " + rangeValues.toString());
  for (var i = 0; i < rangeValues.length; i++) {
    //log("rangeValues["+i+"]["+column+"] [" + rangeValues[i][column] + "] searchCriteria [" + searchCriteria + "]" + " bool " + (rangeValues[i][column] == searchCriteria) + " typeof rangeval " + typeof rangeValues[i][column] + " -- typeof search " + typeof searchCriteria);
    if (rangeValues[i][column] === searchCriteria) {
      results.push(i + 1);
    }
  }
  //log("results " + results.toString() + " -- len " + results.length);
  if (results.length == 0) throw new Error("visualSheet format incorrect, can't find " + searchCriteria + " in range [" + rangeText + "].");

  return results;
}

//Function to control logging with a single global variable for productionization vs debugging
function log(message) {
  if (logging) Logger.log(message);
}

/* DEALING WITH FORM FIELD UPDATES */

var formId = "1-rPqVRS9dulwN4Rirk0VFauI-kAjkX8r1sFQVg_uoWI";
var form = FormApp.openById(formId);

function updateRoomField() {

  var roomSheetId = 1310815220;
  var roomSheet = getSheetById(roomSheetId);

  updateFormField("Camera", roomSheet);
}

//NOTE IF DOING EVERYTHING VIA SHEET, CAN KILL OFF ALL THIS CODE + FORM
//TODO have this pick up the customers
//TODO have a link in the top description to another form to add new customers
//FYI modified contatti from text field to multiple choice to handle customers, let's see if it breaks anything
//TODO for sposta, when sheet update, if column status to sposta, then createEvent from line selected, set status to sposta, force visual update for that line to delete, make a copy of the line to the end and run formsubmit on the new event again
//TODO consider column for "guest" vs "main" in customer profile
  //TODO then add columns for "main" and "guest" up to 3 columns as max guest per room is 4
function updateCustomerField() {

  var roomSheetId = 1310815220;
  var roomSheet = getSheetById(roomSheetId);

  updateFormField("Camera", roomSheet);
}

function updateFormField(fieldTitle, sheet) {
  
  //Get the items from the form and their titles into an array
  var items = form.getItems();
  var titles = items.map(function(item) {
    return item.getTitle();
  })

  //Check for the current title in the array
  var index = titles.indexOf(fieldTitle);
  if (index == -1) throw new Error ("Can't find the question title is in the list of questions " + titles.toString());

  //Get the item to update with the new list of values
  var item = items[index];
  //var itemId = item.getId();

  var values = [];
  values = getListOfValues(fieldTitle, sheet);
  //log("values: " + values.toString());

  updateDropdown(item, values);
}

function getListOfValues(fieldTitle, sheet) {

  var values = [];

  if (sheet) {
    var dataRange = sheet.getDataRange();
    var data = dataRange.getValues()
    for (var i = 2; i < dataRange.getLastRow(); i++) { //start from 2 to skip the first two title rows
      //log("data[i][1] " + data[i][0]);
      if (data[i][0] !== "") {
        if (fieldTitle === "Camera") {
          //log("in here " + data[i][0] + " - " + data[i][1]);
          //Get the category-roomName, to allow for sorting as desired
          values.push(data[i][0] + "-" + data[i][1]);
        } else if (fieldTitle === "Ospiti") {
          values.push(data[i][0]);
          //TODO correct row/col
        }
        
      }
    }
    log("data for " + fieldTitle + "--> " + values.toString());
    if (fieldTitle === "Camera") {
      //This sort works based on the category number appended at the start
      var sorted =  values.sort()

      //Then remove the category number to just have the room names
      return sorted.map(function (v) {
        return v.substring(v.indexOf("-") + 1, v.length);
      })
    } else {
      return values.sort();
    }
  } else {
    throw new Error ("sheet is undefined");
  }
  
}

//Returns a sheet by ID from a given spreadsheet
function getSheetById(id) {
  return SpreadsheetApp.getActive().getSheets().filter(
    function(s) {return s.getSheetId() === id;}
  )[0];
} 

function updateDropdown(item, values) {
  //var item = form.getItemById(itemId);
  item.asListItem().setChoiceValues(values);
}

/* ON EDIT FUNCTIONS
    THESE FUNCTIONS ARE USED TO MODIFY THE BOOKING / VISUAL SPREADSHEETS BASED ON EDITS

    TODO have a feeling that this onEdit will totally exceed daily quotas on Google Drive
*/

function onBookingEdit(event) {

  /*
  var mock = true;
  var event = {
      "authMode":"FULL",
      "oldValue":"chi sa",
      "range":{"columnEnd":10,"columnStart":10,"rowEnd":10,"rowStart":10},
      "source":{},
      "triggerUid":"8038011",
      "user":{"email":"nabeeliphone@gmail.com","nickname":"nabeeliphone"},
      "value":"mondo"
    };
  */

  log('event ' + JSON.stringify(event)); 

  //Get event values
  var range = event.range;
  var oldValue = event.oldValue
  if (oldValue !== null && oldValue !== "") oldValue = oldValue.toString().toLowerCase();

  //log("oldvalue !== null " + (oldValue !== null) + " oldValue !== '' " + (oldValue !== "") + " both cond " + (oldValue !== null && oldValue !== "") + " -- oldValue.toString.lower " + oldValue.toString().toLowerCase() + " -- oldValue lower " + oldValue.toLowerCase());

  var newValue = event.value;
  if(newValue !== null && newValue !== "") newValue = newValue.toString().toLowerCase();

  log("old value (lowercase) " + oldValue);
  log("value (lowercase) " + newValue);

  //TODO store last person who edited the row in the last column
  log("user " + event.user);

  var rangeCol = range.getColumn();
  var rangeRow = range.getRow();
  
  //log("rangeRow " + rangeRow);
  //log("range Col type " + (typeof rangeCol) + " col: " + rangeCol + " numcols " + range.getNumColumns() + " numrows " + range.getNumRows());
  //If more than one range or column are affected at once, ignore cause can't handle multiple bookings in one go
  if (range.getNumColumns() > 1 || range.getNumRows() > 1) {
    log("ignoring as multiple cell changes");
    return;

    //if editing only one row/booking then...
  } else {

    //Get the appropriate sheet
    var ss = event.source;;
    var sheet = ss.getSheetByName("Bookings");

    //Get the whole row of the booking that is being edited
    var bookingRange = sheet.getRange(rangeRow, 1, 1, 20); //20 is the number of columns there are no till COL T
    var bookingValues = bookingRange.getValues();
    
    //Get the cell with the booking ID of the line being edited
    var bookingIdCell = sheet.getRange(rangeRow, 19);
    var bookingId = bookingIdCell.getValues()[0][0]; //19 col for booking id

    //Set the timestamp to show latest update
    var bookingTimestamp = sheet.getRange(rangeRow, 1);
    bookingTimestamp.setValue(new Date());

    //Get the main booking data required
    var camera = bookingValues[0][1];
    var arrival = bookingValues[0][2];
    var departure = bookingValues[0][3];
    var bookingStatus = bookingValues[0][9];

    log("bookingId [" + bookingId + "] -- camera " + camera + " arrival " + arrival + " departure " + departure + " bookingStatus " + bookingStatus);

    //ignore any edits in the header row
    if (rangeRow == 1) {
      log("ignoring edit in headers");
      //if a booking id exists, then editting an existing booking
    } else if (bookingId) {
      log("editing existing booking");
      //TODO if modified range is arrival, departure, camera or status, then re-draw
      //otherwise no action required from this script, details of booking

      //Get the title of the current column
      var titleValue = getHeaderTitle(rangeCol, sheet);

      if (titleValue === "BookingStatus") {
        //TODO need to figure out all the permutations that could make a difference
        //anything to confirmed - then update visual
        //anything to da pagare - then update visual
        //if old value is confirmed or da pagare to anything else - then remove
        var event = loadEventFromBookingSheetRow(bookingRange);

        decideActionBasedOnValueChange(oldValue, newValue, event, rangeRow);
      }
      //if no booking id, then process new booking data if complete
    } else {
      //if required booking data is complete, add new booking
      if (camera && arrival && departure && bookingStatus) {
        //log("adding new booking");

        //Get the nextBookingID value from the counter cell
        var nextBookingIdRange = sheet.getRange(1, 24);
        var nextBookingId = nextBookingIdRange.getValues()[0][0];
        log("nextbookingID: " + nextBookingId);

        //Update this rows booking ID to the correct nextId
        bookingIdCell.setValue("B" + nextBookingId);

        //Update the nextBookingID counter cell with the next value
        var idAsNumber = parseInt(nextBookingId) + 1;
        var idAsString = idAsNumber.toString()
        var lenOfId = idAsString.length;
        nextBookingIdRange.setValue("'" + "0".repeat(6-lenOfId) + idAsNumber);
        //log("idAsNumber " + idAsNumber + " idAsString " + idAsString + " lenOfId " + lenOfId + " -- setting to : " + "0".repeat(6-lenOfId) + idAsNumber);

        var event = loadEventFromBookingSheetRow(bookingRange);

        onFormSubmit(event, rangeRow);
        //if new row edited but doesn't have camera, arrival, departure and booking status, ignore it
      } else {
        log("new row, but doesn't have the required data in it yet");
      } 
    }
  }
}

function test() {
  //decideActionBasedOnValueChange("confermato", "confermato");
  //decideActionBasedOnValueChange("confermato", "da pagare");
  //decideActionBasedOnValueChange("da pagare", "confermato");
  //decideActionBasedOnValueChange("da pagare", "annullato");
  //decideActionBasedOnValueChange("confermato", "sposta");
  //decideActionBasedOnValueChange("sposta", "confermato");
  decideActionBasedOnValueChange("annullato", "da pagare");
}

function decideActionBasedOnValueChange(oldValue, newValue, event, rangeRow) {
  log("in decide: oldval " + oldValue + " new val: " + newValue + " rangeRow " + rangeRow);
  if (oldValue == newValue) {
    //log("ignoring update where old and new values are the same");
    return;
  }

  //if switching between confermato and da pagare, but not the same, then erase and re-draw
  if ((oldValue == "confermato" || oldValue == "da pagare") && (newValue == "confermato" || newValue == "da pagare")) {
    log("erase old visual, and create new one");
    //TODO
    return;
  } 
  
  if ((oldValue == "confermato" || oldValue == "da pagare") && (newValue != "confermato" && newValue != "da pagare")) {
    log("erase old visual only");
    //TODO
    return;
  }

  if ((oldValue != "confermato" && oldValue != "da pagare") && (newValue == "confermato" || newValue == "da pagare")) {
    //log("create new visual only");
    //success
    onFormSubmit(event, rangeRow);
    return;
  }
}

function loadEventFromBookingSheetRow(bookingRange) {

  var bookingValues = bookingRange.getValues();
  //log("bookingRange " + bookingRange + " -- bookingValues " + JSON.stringify(bookingValues) + " -- typeof bookingVals: " + (typeof bookingValues) + " -- len bookingval[0] " + bookingValues[0].length + " -- bookingVal[0] " + bookingValues[0]);
  var event = {
    Timestamp: bookingValues[0][0], 
    Camera: bookingValues[0][1],
    DataArrivo: bookingValues[0][2], 
    DataPartenza: bookingValues[0][3], 
    Acconto: bookingValues[0][5],
    Voucher: bookingValues[0][6],  
    Metodo: bookingValues[0][7], 
    Contatti: bookingValues[0][8], 
    BookingStatus: bookingValues[0][9], 
    Note: bookingValues[0][10], 
    Massagi: bookingValues[0][11], 
    Aperitivo: bookingValues[0][12],
    Dessert: bookingValues[0][13], 
    CestoBio: bookingValues[0][14], 
    Ebike: bookingValues[0][15], 
    Email: bookingValues[0][16], 
    Telefono: bookingValues[0][17]
  }

  //log("event at end of load : " + JSON.stringify(event));
  return event;
}

//Get the header title for the specified column
function getHeaderTitle(column, sheet) {
    var row = 1;
    var titleRange = sheet.getRange(row, column);
    var titleValue = titleRange.getValues()[0][0];
    log("titleValue " + titleValue);

    return titleValue;
}