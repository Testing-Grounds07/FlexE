//Interface methods
function createInterface() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Document')
      .addItem('New', 'createNewSpreadsheet')
      .addItem('Update', 'updateDocument')
      .addSeparator()
      .addItem('Remove', 'showRemovePrompt')
      .addSeparator()
      .addItem('Tags & Symbols', 'openSymbols')
      .addToUi();
}

function openSymbols() {
  var html = HtmlService.createHtmlOutputFromFile('Symbols').setWidth(200).setTitle("Tags & Symbols");
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(html);
}


function showRemovePrompt() {
  var html = HtmlService.createHtmlOutputFromFile('Remove')
      .setWidth(400)
      .setHeight(200);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(html, 'Remove Menu');
}

function removeController(values) {
  
  if(values[0] == "delete")
    remove("document");
  
  if(values[1] == "delete")
    remove("spreadsheet");
  else if(values[1] == "archive")
    archive();
  
}
