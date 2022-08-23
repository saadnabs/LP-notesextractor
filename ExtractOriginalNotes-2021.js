var resultStartRow = 2;
//Result variables  
var resultStartColumn = "A";
var resultEndColumn = "S";
var resultRowWrite = 2;

var bookingsProcessed = 0;

var allRooms = ["Suite Bleue", "Black Cabin", "Bambù", "Quercia", "Rose", "Limoni", "Uva", "More", "Lavanda", "Olivo", "Edera", "Papiro"];

var output;

function main() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  //Explicitly get the search notes sheet called output
  var outputSheetName = "output-test";//"output-originals-21";
  output = spreadsheet.getSheetByName(outputSheetName);

  //clearOutOldResults(output);

  resultStartRow = getLastRow(output) + 1;
  resultRowWrite = resultStartRow;

  var sheets = spreadsheet.getSheets();

  //for (var i in sheets) {
  //  var sheetToSearch = sheets[i]; //sheet // 
  //  var sheetName = sheetToSearch.getName();

    var sheetName =
    "apr-giu 21";
    //"gen-mar 21";
    //"ott-dic 21";
    //"lu-sett 21";
    var sheetToSearch = spreadsheet.getSheetByName(sheetName);

    //log("spreadsheet name: " + sheetName + " -- processing: " + ((!sheetName.includes("output")) && (!sheetName.includes("new ")) && (!sheetName.includes("Form ")) && (!sheetName.includes("Room "))));
    if (!sheetName.includes("output") && !sheetName.includes("new") && !sheetName.includes("done") && !sheetName.includes("Form ") && !sheetName.includes("Room ")) {
      var dateYear = sheetName.slice(-2);
      log("Processing bookings from " + sheetName);

      processRange(sheetToSearch, "A2:AF14", "A1", dateYear);
      processRange(sheetToSearch, "A16:AF28", "A15", dateYear);
      processRange(sheetToSearch, "A30:AF42", "A29", dateYear);
    } // end if not output
  //}  

}

function processRange(sheetToSearch, rangeForData, rangeForMonth, dateYear) {
  //Get the entire data range for the first month section and all the notes    
  var dataRange = sheetToSearch.getRange(rangeForData)
  var dataValues = dataRange.getValues();
  var dateMonth = sheetToSearch.getRange(rangeForMonth).getValue();
  if (dateMonth < 10) { dateMonth = "0" + dateMonth; }
  //log("dateMonth: " + dateMonth);
  var notes = dataRange.getNotes();

  log("Assessing bookings:");
  assessBookingEntries(dataRange, dataValues, notes, dateYear, dateMonth);
}

function assessBookingEntries(dataRange, dataValues, notes, dateYear, dateMonth) {
  for (var i = 0; i < dataValues.length - 1; i++) {

    for (var j = 1; j < dataValues[i].length; j++) {

      var cellValue = dataValues[i][j];
      var dateDay = (j > 9) ? (j) : "0" + (j);
      var dateText = "20" + dateYear + "-" + dateMonth + "-" + dateDay;
      
      //When a booking is found in standard row of a room
      //TODO eventually have S in 2022 when a SPA booking exists
      if (cellValue === "x") {
        //log("found x [" + cellValue + "] in cell at [" + i + ":" + j + "]");

        var room = dataValues[i][0];
        var note = notes[i][j].toLowerCase();

        //log(bookingsProcessed + " room " + room + " -- note: " + note);
        if (note) {
          writeBooking(room, dateText, note);

          //else when there's an X without a note, it's because it's the continuation of a previous booking
          //One example where this breaks, jan 30 bambu booking till jan 31, but then jumps to next X in Limoni that is part of a limoni/bambu booking, and gets counted against the bambu again.
        } else {

          var previousCellValue = dataValues[i][j-1];

          if (previousCellValue === "") {
            //This is not a continuation of the previous cell booking, might be visually connected - 2 rooms booked by 1 guest
            writeBooking(room, dateText, "TBC multi-room booking");
          } else if (allRooms.includes(previousCellValue)){
            writeBooking(room, dateText, "TBC room booking overflowing from last month");
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
              writeBooking(room, dateText, "no note")
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
            if (room == "") {
              log(bookingsProcessed + " -- note: " + note);
              writeBooking("N/A", dateText, note);
            }
          }

          //note = note.substring(note.indexOf(" ") + 1, note.length - 1);
          //log("room " + room + " -- note: " + note);
          writeBooking(room, dateText, note);
        }
        //if cell value is empty, skip
      } else {
        log("shouldn't encounter this");
      }

    }
  }
}

function writeBooking(room, dateText, note) {

  var toOutput = [dateText, room, note];
  writeOutput(resultRowWrite, toOutput);
  log("  booking " + (bookingsProcessed + 1));
  resultRowWrite++;
  bookingsProcessed++;
}

function writeOutput(row, toOutputArray) {

  var startCol = 1;
  var cell = "";

  for (var i = 0; i < toOutputArray.length; i++) {
    cell = output.getRange(row, startCol++);
    cell.setValue(toOutputArray[i]);
  }

  //TODO remove this, it's just for testing to see each line output as it goes rather than batched
  SpreadsheetApp.flush();
}

function clearOutOldResults(outputSheet) {
  var dataRange = outputSheet.getDataRange();
  var lastRow = dataRange.getLastRow();
  var range = outputSheet.getRange(resultStartColumn + resultStartRow + ":" + resultEndColumn + lastRow);
  range.clearContent();
}

function getLastRow(sheet) {
  var dataRange = sheet.getDataRange();
  return dataRange.getLastRow();
}