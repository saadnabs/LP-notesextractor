//NOTES
/*
  VERIFIED OUTPUT FROM OTT->DIC OF '22
  ALL NOTES EXTRACTED APPROPRIATELY AND MULTI DAY BOOKINGS HANDLED CORRECTLY
  NEED TO MANUALLY HANDLE MULTIROOM BOOKINGS AND BOOKINGS OVERFLOWING ACROSS MONTHS

  RUN THIS SCRIPT TO EXTRACT THE CONTENTS FROM THE COMMENTS
  THEN RUN THE SPLIT NOTES SCRIPT ON THE SAME SHEET
*/

var resultRowWrite = 2;

var bookingsProcessed = 0;

var allRooms = ["Suite Bleue", "Black Cabin", "Bambù", "Quercia", "Rose", "Limoni", "Uva", "More", "Lavanda", "Olivo", "Edera", "Papiro"];

var output;
var spreadsheetId;
var previousCellColor = "";
var sheetName;

var outputSheetName = "output-test";

function main22() {
  
  //Get the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  log('ss id ' + spreadsheetId);

  //Get the output sheet
  output = spreadsheet.getSheetByName(outputSheetName);

  //Clear out aold results
  clearOutOldResults(output);

  //If not clearing out results, set the first row to be the last row of the current data range to append
  resultStartRow = getLastRow(output) + 1;
  resultRowWrite = resultStartRow;

  var sheets = spreadsheet.getSheets();

  //for (var i in sheets) {
  //  var sheetToSearch = sheets[i]; //sheet // 
  //  var sheetName = sheetToSearch.getName();

    sheetName =
    "ott-dic 22";
    //"gen-mar 21";
    //"ott-dic 21";
    //"apr-giu 21";
    var sheetToSearch = spreadsheet.getSheetByName(sheetName);

    //log("spreadsheet name: " + sheetName + " -- processing: " + ((!sheetName.includes("output")) && (!sheetName.includes("new ")) && (!sheetName.includes("Form ")) && (!sheetName.includes("Room "))));
    if (!sheetName.includes("output") && !sheetName.includes("new") && !sheetName.includes("done") && !sheetName.includes("Form ") && !sheetName.includes("Room ")) {
      var dateYear = sheetName.slice(-2);
      log("Processing bookings from " + sheetName);

      processRange22(sheetToSearch, "A2:AF16", "A1", dateYear);
      processRange22(sheetToSearch, "A18:AF32", "A17", dateYear);
      processRange22(sheetToSearch, "A34:AF48", "A33", dateYear);
    } // end if not output
  //}  

}

function processRange22(sheetToSearch, rangeForData, rangeForMonth, dateYear) {
  //Get the entire data range for the first month section and all the notes    
  var dataRange = sheetToSearch.getRange(rangeForData)
  var dataValues = dataRange.getValues();
  var dateMonth = sheetToSearch.getRange(rangeForMonth).getValue();
  if (dateMonth < 10) { dateMonth = "0" + dateMonth; }
  //log("dateMonth: " + dateMonth);
  var notes = dataRange.getNotes();

  log("Assessing bookings:");
  assessBookingEntries22(sheetToSearch, dataValues, notes, dateYear, dateMonth);
}

function assessBookingEntries22(sheetToSearch, dataValues, notes, dateYear, dateMonth) {
  //TODO ROW: for testing i = 0
  for (var i = 0; i < dataValues.length - 1; i++) {

    //TODO COLUMN: for testing j = 1 
    for (var j = 1; j < dataValues[i].length; j++) {

      var cellValue = dataValues[i][j];
      var dateDay = (j > 9) ? (j) : "0" + (j);
      var dateText = "20" + dateYear + "-" + dateMonth + "-" + dateDay;
      
      //When a booking is found in standard row of a room
      //TODO eventually have S in 2022 when a SPA booking exists
      if (cellValue === "x" || cellValue === "S"  || cellValue === "T") {
        //log("found x [" + cellValue + "] in cell at [" + i + ":" + j + "]");

        var room = dataValues[i][0];
        var note = notes[i][j].toLowerCase();

        //TODO get borders
        //i+1 cause it's relative to the chosen range, which starts from row 2
        /*
        var rangeOfCell = sheetToSearch.getRange(i+2,j+1,1,1);
        log("range a1not " + (sheetName + "!" + rangeOfCell.getA1Notation()));
        var borders = Sheets.Spreadsheets.get(spreadsheetId, {ranges: sheetName + "!" + rangeOfCell.getA1Notation(), fields: "sheets/data/rowData/values/userEnteredFormat/borders"});
        
        log("borders: " + JSON.stringify(borders));
        var bordersJson = JSON.parse(borders);
        var data = borders.sheets[0].data[0];
        log("data: " + data);

        bordersString = "";

        //borders binary top, right, bottom, left
        if (data == undefined) {
          borders = "0000";
        } else {
          var bs = data.rowData[0].values[0].userEnteredFormat.borders;
          bordersString += checkBorder (bs.top);
          bordersString += checkBorder (bs.right);
          bordersString += checkBorder (bs.bottom);
          bordersString += checkBorder (bs.left);
          
        }
        log("bordersString: " + bordersString);
        */

        //log(bookingsProcessed + " room " + room + " -- note: " + note);
        if (note) {
          writeBooking22(room, dateText, note); //, typeOfBooking(bordersString)

          //else when there's an X without a note, it's because it's the continuation of a previous booking
          //One example where this breaks, jan 30 bambu booking till jan 31, but then jumps to next X in Limoni that is part of a limoni/bambu booking, and gets counted against the bambu again.
        } else {

          var previousCellValue = dataValues[i][j-1];
          
          /*var rangeOfTwoCells = sheetToSearch.getRange(i,j-1,1,2);
          log("range a1not " + rangeOfTwoCells.getA1Notation());
          log("range row " + rangeOfTwoCells.getRow());
          log("range col " + rangeOfTwoCells.getColumn());
          log("range numcols " + rangeOfTwoCells.getNumColumns());
          var currentCellColor = rangeOfTwoCells.getCell(rangeOfTwoCells.getRow(), rangeOfTwoCells.getNumColumns()).getBackground();
          previousCellColor = rangeOfTwoCells.getCell(rangeOfTwoCells.getRow(),rangeOfTwoCells.getColumn()).getBackground();

          log("current color: " + currentCellColor + " -- previousCellColor " + previousCellColor);
          */

          if (previousCellValue === "") {
            //This is not a continuation of the previous cell booking, might be visually connected - 2 rooms booked by 1 guest
            writeBooking22(room, dateText, "TBC multi-room booking");
          } else if (allRooms.includes(previousCellValue)){
            writeBooking22(room, dateText, "TBC room booking overflowing from last month");
          } else {

            var lastRow = getLastRow(output); //processing current one
            var lastBooking = bookingsProcessed;

            log("  booking continutation " + lastBooking);
            var nggCell = output.getRange(lastRow, 5);
            var ngg = nggCell.getValue();
            if (!ngg) { ngg = 1; }
            
            var bookingDate = output.getRange(lastRow, 1).getValue();
            var lastDay = new Date(bookingDate.getFullYear(), bookingDate.getMonth() + 1, 0).getDate();

            if ((bookingDate.getDate() + ngg - 1) <= lastDay) {
              //log(bookingsProcessed + " assessBookingEntries -- contatti: " + output.getRange(lastRow, 7).getValue());;
              //log("assessBookingEntries: last row's ngg: " + ngg);
              nggCell.setValue(++ngg);
              //log("assessBookingEntries: last row's updated ngg: " + ngg);
              //If greater than 31 and has X but no note, then create an entry to assess manually
            } else {
              writeBooking22(room, dateText, "no note")
            }

            //update leaving date to be booking date + ngg
            var currentDate = bookingDate.getDate();

            var leavingDate = new Date(+bookingDate);
            leavingDate.setDate(currentDate + ngg);

            //log("in assess: bookingDate: " + bookingDate + " -- currentDate " + currentDate + " --currentDate + ngg " + (currentDate + ngg) + " leavingDate " + leavingDate);

            //output leaving date based on latest ngg
            output.getRange(lastRow, 4).setValue(leavingDate);
          }
        }


        //When a booking cell is empty in any row
      } else if (cellValue === "") {
        //log("skipping empty cell at [" + i + ":" + j + "]");

        //When a booking is found in the calendar days or dates rows
      } else if (!isNaN(cellValue) || cellValue !== "x") {
        //log("found non-x [" + cellValue + "] in cell at [" + i + ":" + j + "]");
        var note = notes[i][j];

        if (note) {
          //In the calendar days, room is always the first word
          var room = note.substring(0, note.indexOf(" ") - 1).toLowerCase();
          //log(bookingsProcessed + " room: " + room);

          //Fixing specific error where room name isn't added in day/date notes
          if (room === "ok" || room === "ex") {
            room = "Unspecified";
          } else {
            room = getFullRoomName(room);
          }

          //note = note.substring(note.indexOf(" ") + 1, note.length - 1);
          //log("room " + room + " -- note: " + note);
          if (room == "") {
            log(bookingsProcessed + " -- note: " + note);
            writeBooking22("N/A", dateText, note);
            
          } else {
            writeBooking22(room, dateText, note);
          }
        }
        //if cell value is any other unexpected value
      } else {
        writeBooking22("N/A - cell value: " + cellValue, dateText, note);
      }

    }
  }
}

function checkBorder (side) {
  if (side != undefined) {
    var colorStyle = side.colorStyle;
    if (colorStyle != undefined) {
      var rgb = colorStyle.rgbColor;
      if (rgb != undefined) {
        return "1";
      }
    }
  }

  return "0";
          
}

function typeOfBooking(bs) {
  if (bs == "1101" || bs == "0111" || bs == "0101")
    return "Multi-room booking";
  else if (bs == "1011" || bs == "1010" || bs == "1011")
    return "Multi-day booking";
}

function writeBooking22(room, dateText, note, typeOfBooking) {

  var toOutput = [dateText, room, note, typeOfBooking];
  writeOutput22(resultRowWrite, toOutput);
  log("  booking " + (bookingsProcessed + 1));
  resultRowWrite++;
  bookingsProcessed++;
}

function writeOutput22(row, toOutputArray) {

  var startCol = 1;
  var cell = "";

  for (var i = 0; i < toOutputArray.length; i++) {
    cell = output.getRange(row, startCol++);
    cell.setValue(toOutputArray[i]);
  }

  //TODO remove this, it's just for testing to see each line output as it goes rather than batched
  SpreadsheetApp.flush();
}