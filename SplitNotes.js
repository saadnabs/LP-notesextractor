//When taking Cinzia file, replace "Alloggi" with the month number
//Her default is to have the year in YY at the end of the sheet name, this is used in processing

//TODO Sheet ranges in extractNotes depend on number of rooms / will be different over the years

//BE AWARE
//When adding columns
//- modify the variable 'splitEndColumn' definition below
//- update two instances of "toOutput" in the correct order of the additional column

//When taking Cinzia file, replace "Alloggi" with the month number
//Her default is to have the year in YY at the end of the sheet name, this is used in processing

//TODO Sheet ranges in extractNotes depend on number of rooms / will be different over the years

//BE AWARE
//When adding columns
//- modify the variable 'splitEndColumn' definition below
//- update two instances of "toOutput" in the correct order of the additional column

// CURRENT BUGS
//Line 2:
//Nota annullamento is missing the last word for some reason
//num of days not a valid number

var logging = true;
var timing = false;

//split variables  
var splitRowWrite = 2;

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
  var outputSheetName = "split-ott-dic 22";
  output = spreadsheet.getSheetByName(outputSheetName);

  //Clear out previous results
  clearOutOldResults(output);
  
  //Append to sheet from the last row
  var dataRange = output.getDataRange();
  splitStartRow = dataRange.getLastRow() + 1;

  //log("splitstartrow" + splitStartRow);
  
  //log("spreadsheet name: " + spreadsheet.getName());
  var sheets = spreadsheet.getSheets();

  //log("sheets has " + sheets.length + " sheets: " + spreadsheet.getSheets());
  //Cycle through all the sheets

// Force manual processing of sheets, too many bookings, times out
//  for (var i in sheets) {
//    var sheetToSearch = sheets[i]; //sheet // 
//    var sheetName = sheetToSearch.getName();
    var sheetName = "extracted-ott-dic 22";
    var sheetToSearch = spreadsheet.getSheetByName(sheetName);
    log("spreadsheet name: " + sheetName + " -- processing: " + ((!sheetName.includes("output")) && (!sheetName.includes("new ")) && (!sheetName.includes("Form ")) && (!sheetName.includes("Room "))));
    if ((!sheetName.includes("output")) && (!sheetName.includes("new ")) && (!sheetName.includes("Form ")) && (!sheetName.includes("Room "))) {
      var dateYear = sheetName.slice(-2);
      //log("Searching [i]: " + i + " -- " + sheetName);
      log("Processing bookings from " + sheetName);

      var dataRange = sheetToSearch.getDataRange();
      var dataValues = dataRange.getValues();

      //Loop over rows (skipping header), and process the note column
      for (var i = 1; i < dataValues.length; i++) {

        var checkIn = new Date(dataValues[i][0]);
        var room = dataValues[i][1];
        var note = dataValues[i][2];
        var checkOut = dataValues[i][3];
        var numOfDays = dataValues[i][4];

        //TODO dormono OR vedi - leaving for now, adds complexity of multiple entries for one booking

        //% 10 == 0
        if (i == 99)
          log('debug');

        if (!numOfDays && !checkOut) {
          numOfDays = 1;
          checkOut = new Date(+checkIn);
          checkOut.setDate(checkIn.getDate() + numOfDays);
        }

        var toOutput = splitNote(note);
        toOutput.unshift(checkIn, room, checkOut, numOfDays);

        writeOutput(splitRowWrite, toOutput);

        splitRowWrite++;
        bookingsProcessed++;

        if (bookingsProcessed % 25 === 0) {
          log("Processed " + bookingsProcessed + " bookings till now...");
        }

      }
    } // end if not output
// End for
//  }

  //check if sort works without moving header
  //TODO, put this back once done testing lines
  //sortResults(output);

  var endTime = Date.now();
  log("Bookings processed: " + bookingsProcessed + " in " + Math.round(((endTime - startTime) / 1000)) + " seconds.");
}

function sortResults(output) {
  var lastRow = output.getDataRange().getLastRow();
  //TODO use getDataRange
  output.getRange("A2:" + resultEndColumn + lastRow).sort(1);
}
//TODO the splitting process to be cleaned up as I go through different booking formats
function splitNote(note) {

  var pagato = "";
  var paga = "";
  var regalo = "";
  var metodo = "";
  var contatti = "";
  var email = "";
  var telephone = "";
  var nomi = "";
  var status = "";
  var massage = "";
  var massageCount = 0;
  var voucher = "";
  var apertivo = "";
  var costoAperitivo = 0;
  var dessert = "";
  var cestoBio = "";
  var ebike = "";
  var spa = "";
  var fnb = "";
  var fiori = "";
  var ngg = 1; //numero giorni
  var comingFrom = "";
  var dataPagata = "";
  var dataPrenotata = "";
  var noteAnnulla = "", noteSposta = "";
  var originalNote = note;

  //Check if the note starts with the room name, update room name and remove it to process note as usual

  //TODO handle notes that say vedi or rooms N/A or review

  //Standardise bambù to bambu
  if (note.includes("ù")) {
    note.replace("ù", "u");
  }

  //Work on the first line

  //Check for room name at the start of a line - usually for cancelled entries, mostly seen
  if (note.toLowerCase().startsWith("black") || note.toLowerCase().startsWith("bambu") || note.toLowerCase().startsWith("bleue") || note.toLowerCase().startsWith("edera") || note.toLowerCase().startsWith("papiro")) {
    var nextSpaceLocation = regexIndexOf(note, /\s/);
    
    //remove the processed part of the note
    note = removeFromNoteFromTo(note, 0, nextSpaceLocation);
  }

  //Check if voucher
  var results = findValueAndExtract(note, /[a-zA-Z]{3,4}-[0-9]{3,}/);
  note = results.note, voucher = results.valueFound;
  if (voucher) {metodo = "voucher"};

  //Check if paid amount present
  //TODO ISSUE with extracting line 12 with pagato 1.050, but extracts ## from date scad
  //only impacts a few entries. can NOT add xero, caause it leaves CCOK attachr

   var results = findValueAndExtract(note, /[0-9.]{2,5}\s/);
  note = results.note, pagato = results.valueFound;
  if (pagato && pagato.includes(".")) {
    pagato = pagato.replace(/\./g,'');
  }

  //Check if paga amount present
  var results = findValueAndExtract(note, /paga [0-9.]{2,5}\s?/, 4);
  note = results.note, paga = results.valueFound;
  if (paga && paga.includes(".")) {
    paga = paga.replace(/\./g,'');
  }

  //Check if regalo amount present
  var results = findValueAndExtract(note, /regalo/);
  note = results.note, regalo = results.valueFound;
  if (regalo && regalo == "regalo") {
    regalo = "si";
  }

  //Check if "ok" status
  var results = findValueAndExtract(note, /ok\s/);
  note = results.note, status = results.valueFound = "ok" ? "confermato" : "";

  //Check if "cc" as payment method
  var results = findValueAndExtract(note, /\s?cc\s/);
  note = results.note;
  if (results.valueFound != undefined) {
    metodo = results.valueFound;
  }  

  //Check if pagat* il is present
  var results = findValueAndExtract(note, /pagat.\sil\s/, 10);
  note = results.note, dataPagata = results.valueFound;

  //Check if aperitivo is within a line, and extract cost
  var results = findValueAndExtract(note, /\+ aperitivo\s*[0-9]{2} \+/, 12);
  note = results.note, costoAperitivo = results.valueFound;
  if (costoAperitivo && costoAperitivo != 0) {
    apertivo = "si";
  }

  //TODO to validate against data
  if (metodo == "" && pagato != "") {
    metodo = "bonifico";
  }

  //Look through the lines of the note to find special lines
  var lines = note.split("\n");
  var removeProcessedLinesFromNote = [];   

  for (var z = 0; z < lines.length; z++) {
    
    // anything with @ or numbers for the contact info 
    //TODO phone numbers will not always be 10 digits, foreign numbers need to see if Ci puts + or 00 and how many numbers. Maybe tel +x numbers
    if ((lines[z].includes("@") || lines[z].match("[0-9]{10}")) && contatti == "") {

      contatti = lines[z];
      //TODO check if these work
      email = contatti.match(/\b[a-zA-Z0-9_\+%.-]+@[a-zA-Z0-9_\+%.-]+\.[a-zA-z]{2,}\b/g);
      if (email && email.length > 0) email = email.join(", ").trim();
      telephone = contatti.match(/(00|\+)?[0-9\s\-\/]{10,}(?![:a-z])/g);
      if (telephone && telephone.length > 0) telephone = telephone.join(", ").trim();
      
      //                      remove dashes     tel text          email                                       phone number
      nomi = contatti.replace(/-/g, "").replace(/tel./g, "").replace(/\b[a-zA-Z0-9_\+%.-]+@[a-zA-Z0-9_\+%.-]+\.[a-zA-z]{2,}\b/g, "").replace(/(00|\+)?[0-9 -\/]{10,}(?![:a-z])/g, "")
      removeProcessedLinesFromNote.push(z);
      continue;
    }

    //check for "prenotat* il" date
    if (lines[z].match(/prenotat.\sil\s/)) {
      var results = findValueAndExtract(note, /prenotat.\sil\s/, 13);
      note = results.note, dataPrenotata = results.valueFound;
      removeProcessedLinesFromNote.push(z);
      continue;
    }

    //check for "annull* il" date
    if (lines[z].match(/annull.\sil\s?/)) {
      //var results = findValueAndExtract(lines[z], /annull.\sil\s?/, 10, lines[z].length);
      //note = results.note, noteAnnulla = results.valueFound;
      noteAnnulla = lines[z];
      status = "Annullato";
      removeProcessedLinesFromNote.push(z);
      continue;
    }

    //check for "spost* il" date
    if (lines[z].match(/spost.*\s?/)) {
      //var results = findValueAndExtract(note, /spost.*\s?/, 10, 1);
      //note = results.note, noteSposta = results.valueFound;
      noteSposta = lines[z];
      removeProcessedLinesFromNote.push(z);
      continue;
    }

    //check for spa,sauna,idro?
    if (lines[z].match(/\+ spa\s?/) || lines[z].match(/\+ sauna\s?/)) {
      nextSpaceLocation = regexIndexOf(lines[z], /\s/);
      spa = lines[z].substring(nextSpaceLocation + 1, lines[z].length);
      removeProcessedLinesFromNote.push(z);
      continue;
    }

    //check for aperitivo in a separate line
    if (lines[z].match(/\+ aperitivo\s.*$\s?/)) {
      if (apertivo == "") {
        apertivo = lines[z].replace(/\+ aperitivo (da )?/, "");
        var costLocation = regexIndexOf(apertivo, /(?<![\/.])[0-9]{2}(?![\/.])/);
        costoAperitivo = apertivo.substring(costLocation, costLocation + 2);
        if (apertivo.length == 2) apertivo = "si";
        
        removeProcessedLinesFromNote.push(z);
        continue;
      } // else
          //if aperitivo is not empty, might have 2 aperitivo lines, leaving it to show in left over note
      
    }

    //check for massaggi/o
    if (lines[z].match(/^\+ [0-9]{1} massaggo?\s?/)) {
      massage = lines[z];
      massageCount = lines[z].match(/[0-9]{1}/);
      removeProcessedLinesFromNote.push(z);
      continue;
    }

    //check for cesto bio
    if (lines[z].match(/^\+ cesto bio\s?/)) {
      cestoBio = "si";
      removeProcessedLinesFromNote.push(z);
      continue;
    }

    //check for cesto bio
    if (lines[z].match(/^\+ fiori\s?/)) {
      fiori = lines[z];
      removeProcessedLinesFromNote.push(z);
      continue;
    }

    //check for menu
    if (lines[z].toLowerCase().match(/menu\s?/) || lines[z].toLowerCase().match(/cena\s?/)) {
      fnb = lines[z];
      removeProcessedLinesFromNote.push(z);
      continue;
    }

    //check for dessert
    if (lines[z].toLowerCase().match(/dessert\s?/)) {
      dessert = lines[z];
      dessert = dessert.replace("+ dessert ", "");
      removeProcessedLinesFromNote.push(z);
      continue;
    }

    //check for dessert
    if (lines[z].toLowerCase().match(/e-bike\s?/)) {
      ebike = lines[z];
      removeProcessedLinesFromNote.push(z);
      continue;
    }

    //check for region
    //if (regions.includes(lines[z].toLowerCase())) {
    var searchRegions = regions.findIndex(element => lines[z].toLowerCase().includes(element))
    if (searchRegions != -1) { 
      if (comingFrom == "") {
        comingFrom = lines[z];
        removeProcessedLinesFromNote.push(z);
        continue;
      }
    }
  }

  //After processing lines, remove the ones extracted
  for (var i = removeProcessedLinesFromNote.length -1; i >= 0; i--)
   lines.splice(removeProcessedLinesFromNote[i],1);

  note = lines.join("\n");

  //RUN1: row 24 has rinforzato , row 7, 12, 31 has normale\
  /* Handled aperitivo above
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
  */

  //TODO maybe line that starts with ref. is for the person who bougth the voucher.. new column

  //Check for status and payment
  //TODO first row not showing correct status
  note = checkStatus(note, status)

  //Check payment methods
  if (note.includes("fa bon")) {
    metodo = "bonifico";
    status = "da pagare";
  } else if (note.includes("salda")) {
    status = "da pagare";
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

  //Check for massages, can be in the format
  //+ [0-9] massagg* ... --> processed as a line in the section before
  //... + massagg ... --> processed here
  var massageAt = note.indexOf("massag");
  if (massageAt != -1) {
    massage = "si";
    massageCount = note.substring(massageAt - 2, massageAt - 1);
    note = note.substring(0, massageAt - 5) + note.substring(massageAt + 8);
    //log("massage count: " + massage);
  }

  //Check for cesto bio
  //TODO might need this if + cesto is inline with toher stuff rather than separate line that is handled above
  /*
  if (note.includes("cesto bio")) {
    cestoBio = "si";
  }

  //Check for bike
  if (note.includes("bike")) {
    ebike = "si";
  }
  */

  return [pagato, dataPagata, paga, dataPrenotata, voucher, regalo, metodo, contatti, nomi, telephone, email, status, massage, massageCount == 0 ? "" : massageCount, fnb, apertivo, costoAperitivo, dessert, cestoBio, fiori, ebike, spa, comingFrom, noteAnnulla, noteSposta, note, originalNote];
  
}

function checkStatus(note, status) {
  var stringToRemove = "";
  if (note.includes("mai pagat")) {
    status = "cancellato";
    stringToRemove = "mai pagat";
  } else if (note.includes("cancellata")) {
    status = "cancellato";
    stringToRemove = "cancellato";
  } else if (note.includes("pagata") || note.includes("pagati")) {
    if (status == "")
      status = "confermato";
    //else
      //not setting it cause it's been set by a previous criteria.
    stringToRemove = "pagat";
  } else if (note.includes("vedi")) {
    status = "manual - referral";
  } else if (status != "") {
    //ignoring status when not empty
  } else {
    status = "altro";
  }

  //RUN1: row 29, note removes mai pagata
  if (stringToRemove != "") {
    var stringToRemoveLocation = note.indexOf(stringToRemove);
    var stringLength = stringToRemove.length;
    if (stringToRemove == "pagat") stringLength++;
    note = note.substring(0, stringToRemoveLocation) + note.substring(stringToRemoveLocation + stringLength, note.length)
  }

  return note;
}

function writeOutput(row, toOutputArray) {

  var startCol = 1;
  var cell = "";

  for (var i = 0; i < toOutputArray.length; i++) {
    cell = output.getRange(row, startCol++);
    cell.setValue(toOutputArray[i]);

  //TODO remove this, it's just for testing to see each line output as it goes rather than batched
  }
  //SpreadsheetApp.flush();

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
  return note.substring(0, startLocation) + note.substring(endLocation + 1, note.length).trim();
}

function findValueAndExtract(note, regexOrString, extraIndex, tillEnd) {

  if(!extraIndex) extraIndex = 0;

  var varLocation = regexIndexOf(note, regexOrString);
  var nextSpaceLocation = regexIndexOf(note, /\s/, varLocation + 1 + extraIndex);

  //Either pass in 1 to go to end of note, or pass in a number to go to a specific end location
  if (tillEnd) {
    if(tillEnd == 1) {
      nextSpaceLocation = note.length;
    } else if (tillEnd > 1) {
      nextSpaceLocation = tillEnd;    
    }
  }

  if (varLocation != -1) {
    var valueFound = note.substring(varLocation + extraIndex, nextSpaceLocation).trim();
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

