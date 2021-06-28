log = console.log

// Add menu

function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Extract YAMM times')
    .addItem('Run', 'run')
    .addToUi();
}

function run() {
  var ui = SpreadsheetApp.getUi();
  var msColumn = ui.prompt("Letter for Merge Status column: ").getResponseText();
  var range_start = ui.prompt("Start row number: ").getResponseText();
  var range_end = ui.prompt("End row number: ").getResponseText();
  var tColumn = ui.prompt("Letter for time column: ").getResponseText();
  addColumnValuesForYAMMTimes(range_start, range_end, msColumn, tColumn)
}

function addColumnValuesForYAMMTimes(low, high, column, columnForTimes) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  for(let i = parseInt(low); i <= parseInt(high); i++) 
  {
    let range = sheet.getRange(column + i);
    let note = range.getNote();
    range = sheet.getRange(columnForTimes + i);
    range.setValue(note);
  }
}

onOpen()