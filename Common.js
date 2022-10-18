//Result variables  
var resultStartRow = 2;
var resultStartColumn = "A";
var resultEndColumn = "AC";

function clearOutOldResults(outputSheet) {
  var dataRange = outputSheet.getDataRange();
  var lastRow = dataRange.getLastRow();

  //Trying to handle the clearing of the headers when the sheet is empty
  if (!resultStartRow || resultStartRow < 2) resultStartRow = 2;
  var range = outputSheet.getRange(resultStartColumn + resultStartRow + ":" + resultEndColumn + lastRow);
  log("in clearOutOldResults -- range " + range.getA1Notation());
  range.clearContent();
}

function getLastRow(sheet) {
  var dataRange = sheet.getDataRange();
  return dataRange.getLastRow();
}