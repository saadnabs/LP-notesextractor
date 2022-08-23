//When taking Cinzia file, replace "Alloggi" with the month number
//Her default is to have the year in YY at the end of the sheet name, this is used in processing

//TODO Sheet ranges in extractNotes depend on number of rooms / will be different over the years
//TODO metodo jan 10 is a number+number

//BE AWARE
//When adding columns
//- modify the variable 'resultEndColumn' definition below
//- update two instances of "toOutput" in the correct order of the additional column

//?? below
//TODOs noteExtraction
//Room name [euro] not found in room list

var logging = true;
var timing = false;

//Result variables  
var resultStartRow = 2;
var resultStartColumn = "A";
var resultEndColumn = "S";
var resultRowWrite = 2;

var rowsBetweenMonths = 0;

var bookingsProcessed = 0;
var bookingValues = "";
var bookingDataRange = "";
var visualDataRange = "";
var visualDataValues = "";

var allRooms = ["Suite Bleue", "Black Cabin", "Bambù", "Quercia", "Rose", "Limoni", "Uva", "More", "Lavanda", "Olivo", "Edera", "Papiro"];

var output;

function extractNotes(outputSheetName) {

  var startTime = Date.now();
  //Get the active spread sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  //Explicitly get the search notes sheet called output
  var outputSheetName = "output-originals";
  output = spreadsheet.getSheetByName(outputSheetName);

  //Clear out previous results
  //clearOutOldResults(output);
  
  //Append to sheet from the last row
  var dataRange = output.getDataRange();
  resultStartRow = dataRange.getLastRow() + 1;

  //log("resultstartrow" + resultStartRow);
  
  //log("spreadsheet name: " + spreadsheet.getName());
  var sheets = spreadsheet.getSheets();

  //log("sheets has " + sheets.length + " sheets: " + spreadsheet.getSheets());
  //Cycle through all the sheets

// Force manual processing of sheets, too many bookings, times out
//  for (var i in sheets) {
//    var sheetToSearch = sheets[i]; //sheet // 
//    var sheetName = sheetToSearch.getName();
    var sheetName = "apr-giu 21";
    var sheetToSearch = spreadsheet.getSheetByName(sheetName);
    log("spreadsheet name: " + sheetName + " -- processing: " + ((!sheetName.includes("output")) && (!sheetName.includes("new ")) && (!sheetName.includes("Form ")) && (!sheetName.includes("Room "))));
    if ((!sheetName.includes("output")) && (!sheetName.includes("new ")) && (!sheetName.includes("Form ")) && (!sheetName.includes("Room "))) {
      var dateYear = sheetName.slice(-2);
      //log("Searching [i]: " + i + " -- " + sheetName);
      log("Processing bookings from " + sheetName);

      processRange(sheetToSearch, "A2:AF14", "A1", dateYear);
      processRange(sheetToSearch, "A16:AF28", "A15", dateYear);
      processRange(sheetToSearch, "A30:AF42", "A29", dateYear);
    } // end if not output
// End for
//  }

  //TODO, check if sort works without moving header
  sortResults(output);

  var endTime = Date.now();
  log("Bookings processed: " + bookingsProcessed + " in " + Math.round(((endTime - startTime) / 1000)) + " seconds.");
}

function sortResults(output) {
  var lastRow = output.getDataRange().getLastRow();
  //TODO use getDataRange
  output.getRange("A1:T" + lastRow).sort(1);
}

function processRange(sheetToSearch, rangeForData, rangeForMonth, dateYear) {
  //Get the entire data range for the first month section and all the notes    
  var dataRange = sheetToSearch.getRange(rangeForData)
  var dataValues = dataRange.getValues();
  var dateMonth = sheetToSearch.getRange(rangeForMonth).getValue();
  if (dateMonth < 10) { dateMonth = "0" + dateMonth; }
  //log("dateMonth: " + dateMonth);
  var notes = dataRange.getNotes();

  assessBookingEntries(dataRange, dataValues, notes, dateYear, dateMonth);
}

function assessBookingEntries(dataRange, dataValues, notes, dateYear, dateMonth) {
  //log("dataRange values: " + dataRange.getValues());
  //Iterate over rows
  rowLoop:
  for (var i = 0; i < dataValues.length - 1; i++) {

    //Iterate over columns
    columnLoop:
    for (var j = 1; j < dataValues[i].length; j++) {

      var cellValue = dataValues[i][j];
      var dateDay = (j > 9) ? (j) : "0" + (j);
      var dateText = "20" + dateYear + "-" + dateMonth + "-" + dateDay;
      //log("isNan " + isNaN(cellValue));

      //When a booking is found in standard row of a room
      //TODO eventually have S in 2022 when a SPA booking exists
      if (cellValue === "x") {
        //log("found x [" + cellValue + "] in cell at [" + i + ":" + j + "]");

        var room = dataValues[i][0];
        var note = notes[i][j].toLowerCase();

        //log(bookingsProcessed + " room " + room + " -- note: " + note);
        if (note) {
          splitNoteAndWriteBooking(room, note, dateText);

          //else when there's an X without a note, it's because it's the continuation of a previous booking
          //One example where this breaks, jan 30 bambu booking till jan 31, but then jumps to next X in Limoni that is part of a limoni/bambu booking, and gets counted against the bambu again.
        } else {

          var lastRow = bookingsProcessed + 1; //processing current one

          //log("assessBookingEntries: last row in output: " + lastRow);
          var nggCell = output.getRange(lastRow, 4);
          var ngg = nggCell.getValue()

          var status = output.getRange(lastRow, 9).getValue();

          var bookingDate = output.getRange(lastRow, 1).getValue();
          var lastDay = new Date(bookingDate.getFullYear(), bookingDate.getMonth() + 1, 0).getDate();

          if ((bookingDate.getDate() + ngg - 1) < lastDay && status != "manual - referral") {
            //log(bookingsProcessed + " assessBookingEntries -- contatti: " + output.getRange(lastRow, 7).getValue());;
            //log("assessBookingEntries: last row's ngg: " + ngg);
            nggCell.setValue(++ngg);
            //log("assessBookingEntries: last row's updated ngg: " + ngg);
            //If greater than 31 and has X but no note, then create an entry to assess manually
          } else {
            writeBookingForReview(room, dateText, "")
          }

          //update leaving date to be booking date + ngg
          var currentDate = bookingDate.getDate();

          var leavingDate = new Date(+bookingDate);
          leavingDate.setDate(currentDate + ngg);

          //log("in assess: bookingDate: " + bookingDate + " -- currentDate " + currentDate + " --currentDate + ngg " + (currentDate + ngg) + " leavingDate " + leavingDate);

          //output leaving date based on latest ngg
          output.getRange(lastRow, 3).setValue(leavingDate);
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
              writeBookingForReview("N/A", dateText, note);
            }
          }

          //note = note.substring(note.indexOf(" ") + 1, note.length - 1);
          //log("room " + room + " -- note: " + note);
          splitNoteAndWriteBooking(room, note, dateText);
        }
        //if cell value is empty, skip
      } else {
        log("shouldn't encounter this");
      }

    }
  }
}
//TODO the splitting process to be cleaned up as I go through different booking formats
function splitNoteAndWriteBooking(room, note, dateText) {

  var pagato = "";
  var metodo = "";
  var contatti = "";
  var nomi = "";
  var status = "";
  var massage = "";
  var voucher = "";
  var apertivo = "";
  var dessert = "";
  var cestoBio = "";
  var ebike = "";
  var ngg = 1; //numero giorni
  var dataPagata = "";
  var dataPrenotata = "";
  var originalNote = note;

  //Check if the note starts with the room name, update room name and remove it to process note as usual

  //Standardise bambù to bambu
  if (note.includes("ù")) {
    note.replace("ù", "u");
  }

  //Work on the first line

  //Check for room name at the start of a line - usually for cancelled entries, mostly seen
  if (note.toLowerCase().startsWith("black") || note.toLowerCase().startsWith("bambu")) {
    var nextSpaceLocation = regexIndexOf(note, /\s/);
    
    room = getFullRoomName(note.substring(0, note.indexOf(" ") - 1));

    //log(bookingsProcessed + " starts with Black or bambu -- substring " + note.substring(0, note.indexOf(" ") - 1) + "-- room " + room + " -- note " + note );
    //remove the processed part of the note
    note = removeFromNoteFromTo(note, 0, nextSpaceLocation);
    //note.substring(note.indexOf(" ") + 1, note.length - 1);
  }

  //Check for paid amount in first line
  /*
  var paidAmountLocation = regexIndexOf(note, /\s(?<!-)[0-9]{3,4}\s/) + 1;
  nextSpaceLocation = regexIndexOf(note, /\s/, paidAmountLocation + 1);

  if (paidAmountLocation != -1) {
    pagato = note.substring(paidAmountLocation, nextSpaceLocation);
    note = removeFromNoteFromTo(note, paidAmountLocation, nextSpaceLocation);
  }
  */

  var results = findValueAndExtract(note, /[a-zA-Z]{3}-[0-9]{3,}/);
  note = results.note, voucher = results.valueFound;
  if (voucher) {metodo = "voucher"};

  var results = findValueAndExtract(note, /[0-9]{3,4}\s/);
  note = results.note, pagato = results.valueFound;

  var results = findValueAndExtract(note, /ok\s/);
  note = results.note, status = results.valueFound = "ok" ? "confermato" : "";

  var results = findValueAndExtract(note, /\s?cc\s/);
  note = results.note, metodo = results.valueFound;

  var results = findValueAndExtract(note, /pagat.\sil\s/, 10);
  note = results.note, dataPagata = results.valueFound;

/*
  var okLocation = regexIndexOf(note, /\sok\s/) + 1;
  nextSpaceLocation = regexIndexOf(note, /\s/, okLocation + 1);

  //Check for ok, to set default status, will check for other trigger words after
  if (okLocation != -1) {
    status = "confermato";
    note = removeFromNoteFromTo(note, okLocation, nextSpaceLocation);
  }
  //RUN1: row 31 check bambu voucher number is 1030 and other not damaged
  //Check for voucher location in first line
  var voucherStartLoc = regexIndexOf(note, /[a-zA-Z]{3}-[0-9]{3,}/);
  nextSpaceLocation = regexIndexOf(note, /\s/, voucherStartLoc + 1);

  if (voucherStartLoc != -1) {
    voucher = note.substring(voucherStartLoc, nextSpaceLocation);
    metodo = "voucher";
    note = removeFromNoteFromTo(note, voucherStartLoc, nextSpaceLocation);
  }

  //Check for cc location in first line
  var ccStartLoc = regexIndexOf(note, /\s?cc\s/);
  nextSpaceLocation = regexIndexOf(note, /\s/, ccStartLoc + 1);

  if (ccStartLoc != -1) {
    metodo = note.substring(ccStartLoc, 2);
    note = removeFromNoteFromTo(note, ccStartLoc, nextSpaceLocation);
  }
*/

  //TODO to validate against data
  if (metodo == "" && pagato != "") {
    metodo = "bonifico";
  }

  //check for "pagat* il" date
  /*
  var pagatailStartLoc = regexIndexOf(note, /pagat.\sil\s/);
  nextSpaceLocation = regexIndexOf(note, /\s/, pagatailStartLoc + 10);

  if (pagatailStartLoc != -1) {  
    //go to end of pagata il (10 chars) and take the next 10 characters
    dataPagata = note.substring(pagatailStartLoc + 10 ,nextSpaceLocation)

    //remove the processed part in the middle of the note, leaving only the rest
    note = removeFromNoteFromTo(note, pagatailStartLoc, nextSpaceLocation);
  }*/

  //Look through the lines of the note to find anything with @ or numbers for the contact info
  var lines = note.split("\n");
  var removeProcessedLinesFromNote = [];    
  for (var z = 0; z < lines.length; z++) {
    //TODO phone numbers will not always be 10 digits, foreign numbers need to see if Ci puts + or 00 and how many numbers. Maybe tel +x numbers
    if ((lines[z].includes("@") || lines[z].match("[0-9]{10}")) && contatti == "") {
      contatti = lines[z];
      //                      remove dashes     tel text          email                                       phone number
      nomi = contatti.replace(/-/g, "").replace(/tel./g, "").replace(/\b[a-zA-Z0-9_\+%.-]+@[a-zA-Z0-9_\+%.-]+\.[a-zA-z]{2,}\b/g, "").replace(/(00|\+)?[0-9 -\/]{10,}(?![:a-z])/g, "")
      removeProcessedLinesFromNote.push(z);

  }

    if (lines[z].match(/prenotat.\sil\s/)) {
      //check for "prenotat* il" date


      var results = findValueAndExtract(note, /prenotat.\sil\s/, 13);
      note = results.note, dataPrenotata = results.valueFound;

      /*
      var prenotatilStartLoc = regexIndexOf(note, /prenotat.\sil\s/);
      nextSpaceLocation = regexIndexOf(note, /\s/, prenotatilStartLoc + 14);

      if (prenotatilStartLoc != -1) {  
        //go to end of pagata il (13 chars) and take the next 10 characters
        dataPrenotata = note.substring(prenotatilStartLoc + 13 ,nextSpaceLocation)
      }
      */

      removeProcessedLinesFromNote.push(z);
  
    }
  }

  //After processing lines, remove the ones extracted
  for (var i = removeProcessedLinesFromNote.length -1; i >= 0; i--)
   lines.splice(removeProcessedLinesFromNote[i],1);

  note = lines.join("\n");

  //RUN1: row 24 has rinforzato , row 7, 12, 31 has normale
  if (note.includes("aperitivo")) {
    //Check for apertifs
      if (note.includes("aperitivo rinforzato")) {
        apertivo = "rinforzato";

        var results = findValueAndExtract(note, "aperitivo rinforzato");
        note = results.note, throwAway = results.valueFound;
      } else if (note.includes("aperitivo express")) {
        apertivo = "express";

        var results = findValueAndExtract(note, "aperitivo express");
        note = results.note, throwAway = results.valueFound;
      } else if (note.includes("aperitivo")) {
        apertivo = "normale";

        var results = findValueAndExtract(note, "aperitivo");
        note = results.note, throwAway = results.valueFound;
      }
  }

  //TODO maybe line that starts with ref. is for the person who bougth the voucher.. new column

  //Check for status and payment
  var stringToRemove = "";
  if (note.includes("sposta ")) {
    status = "sposta";
    stringToRemove = "sposta ";
  } else if (note.includes("annulla")) {
    //TODO check, sometimes C writes annullatto talking about previous cancelation
    status = "annullato";
    stringToRemove = "annulla";
  } else if (note.includes("mai pagat")) {
    status = "cancellato";
    pagato = "N/A";
    metodo = "N/A";
    stringToRemove = "mai pagat";
  } else if (note.includes("cancellata")) {
    status = "cancellato";
    stringToRemove = "cancellato";
  } else if (note.includes("pagata") || note.includes("pagati")) {
    status = "confermato";
    stringToRemove = "pagat";
  } else if (note.includes("vedi")) {
    status = "manual - referral";
  } else {
    status = "altro";
  }

  //RUN1: row 29, note removes mai pagata
  if (stringToRemove != "") {
    var stringToRemoveLocation = note.indexOf(stringToRemove);
    note = note.substring(0, stringToRemoveLocation) + note.substring(stringToRemoveLocation + stringToRemove.length, note.length - 1)
  }

  //Check payment methods
  if (note.includes("fa bon")) {
    metodo = "bonifico";
    status = "da pagare";
  } else if (note.includes("salda")) {
    status = "da pagare";
  } else if (note.includes("cc pagat")) {
    metodo = "carta di credito - pagata";
  } else if (note.includes("garanzia")) {
    metodo = "carta di credito - garanzia";
    status = "confermato";
  } 
  
  else if (note.includes("c-trad-")) {
    var lines = note.split("\n");
    var res = lines[0].split(" ");

    metodo = "voucher";
    voucher = note.substring(note.indexOf("c-trad-") + 7, note.indexOf(" ok "));
    pagato = note.substring(note.indexOf(" ok ") + 4, note.indexOf(" ", note.indexOf(" ok ") + 4));

    /*if(note.startsWith != "c-trad-") {
      pagato = res[3];
      voucher = res[1];   
    } else {
      pagato = res[2];
      voucher = res[0]; 
    }*/

    //TODO remove c-trad text items from resto della nota

  } else if (note.includes("senza caparra")) {
    pagato = 0;
    metodo = "senza caparra";
  } else if (note.includes("salda alla partenza")) {
    metodo = "salda alla partenza";
    pagato = 0;
  }
  else if (metodo == "") {
    metodo = "unspecified";
  }

  //Check for massages
  var massageAt = note.indexOf("massag");
  if (massageAt != -1) {
    massage = note.substring(massageAt - 2, massageAt - 1);
    //log("massage count: " + massage);
  }

  //Check for desserts
  if (note.includes("dessert")) {
    dessert = "si";
  }

  //Check for cesto bio
  if (note.includes("cesto bio")) {
    cestoBio = "si";
  }

  //Check for bike
  if (note.includes("bike")) {
    ebike = "si";
  }

  var bookingDate = new Date(dateText);
  var currentDate = bookingDate.getDate();

  var leavingDate = new Date(+bookingDate);
  leavingDate.setDate(currentDate + ngg);
  //log("in splitNote: bookingDate: " + bookingDate + " -- currentDate " + currentDate + "  --currentDate + ngg " + (currentDate + ngg) + " leavingDate " + leavingDate);

  //log(bookingsProcessed + " contatti: " + contatti + " -- ngg: " + ngg);

  var toOutput = [dateText, room, leavingDate, ngg, pagato, dataPagata, dataPrenotata, voucher, metodo, contatti, nomi, status, massage, apertivo, dessert, cestoBio, ebike, note, originalNote];
  writeOutput(resultRowWrite, toOutput);

  resultRowWrite++;
  bookingsProcessed++;

  if (bookingsProcessed % 25 === 0) {
    log("Processed " + bookingsProcessed + " bookings till now...");
  }
}

function writeBookingForReview(room, dateText, note) {
  //log("outputting entry for x with no note ")
  var toOutput = [dateText, room, "", "", "", "", "", "", "", "", "", "manual - referral", "", "", "", "", "", "", note];
  writeOutput(resultRowWrite, toOutput);

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

//Function to clear all the output and start over
function clearOutOldResults(outputSheet) {
  var dataRange = outputSheet.getDataRange();
  var lastRow = dataRange.getLastRow();
  var range = outputSheet.getRange(resultStartColumn + resultStartRow + ":" + resultEndColumn + lastRow);
  range.clearContent();
}

function titleCase(str) {
  return str.toLowerCase().split(' ').map(function (word) {
    return word.replace(word[0], word[0].toUpperCase());
  }).join(' ');
}

function getFullRoomName(room) {
  var indexOfRoomName = allRooms.findIndex(element => element.includes(titleCase(room)));
  //log(bookingsProcessed + " looking for " + room + " --first 3 letters: " + room.toLowerCase().substring(0, 3));
  if (room.toLowerCase().substring(0, 3) == "bcm") {
    return "Black Cabin";
  } else if (room.toLowerCase().substring(0, 3) == "sbm") {
    return "Suite Bleue";
  } else if (indexOfRoomName != -1) {
    return allRooms[indexOfRoomName];
  } else {
    log(bookingsProcessed + " -- Room name [" + room + "] not found in room list");
    return "";
  }
}

function regexIndexOf(string, regex, startpos) {
  var indexOf = string.substring(startpos || 0).search(regex);
  return (indexOf >= 0) ? (indexOf + (startpos || 0)) : indexOf;
}

function removeFromNoteFromTo(note, startLocation, endLocation) {
  return note.substring(0, startLocation) + note.substring(endLocation + 1, note.length - 1).trim();
}

function findValueAndExtract(note, regexOrString, extraIndex) {

  if(!extraIndex) extraIndex = 0;

  var varLocation = regexIndexOf(note, regexOrString);
  var nextSpaceLocation = regexIndexOf(note, /\s/, varLocation + 1 + extraIndex);

  if (varLocation != -1) {
    var valueFound = note.substring(varLocation + extraIndex, nextSpaceLocation);
    note = removeFromNoteFromTo(note, varLocation, nextSpaceLocation);
  }

  return { note, valueFound };
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('La Piantata')
    .addItem("Extract notes", "extractNotes")
    .addItem("Update Visuals", "updateVisualFromBookingSheet")
    .addToUi();
}

/************************************************** 
//THIS SECTION IS RELATING TO THE ATTEMPT OF A BOOKING SYSTEM NOT NOTE EXTRACTION
/************************************************** */


//TODOs onFormSubmit
//How to handle changes to bookings and updating the visual?
//Validation in form? check if dates are not too far apart? pop up box to accept and continue or to cancel?
//Replace confermato/da pagere free text to static variable, so it's one change everywhere if needed
//Add lookup formula to find customer id
//in customer 
    //add formula to create customer id
    //add formula to count how many stays and how many nights

