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
      .addItem('Help', 'openSymbols')
      .addToUi();
  
  ui.createMenu('Style')
      .addSubMenu(ui.createMenu('Header')
                  .addItem("Title", "styleHeaderTitle")
                  .addItem("Header 1", "styleHeader1")
                  .addItem("Header 2", "styleHeader2")
                  .addItem("Header 3", "styleHeader3")
                  .addItem("Header 4", "styleHeader4")
                  .addItem("Header 5", "styleHeader5")
                  .addItem("Header 6", "styleHeader6"))
      .addSubMenu(ui.createMenu('List Items')
                  .addItem("Nest 1", "styleListItem1")
                  .addItem("Nest 2", "styleListItem2")
                  .addItem("Nest 3", "styleListItem3")
                  .addItem("Nest 4", "styleListItem4")
                  .addItem("Nest 5", "styleListItem5")
                  .addItem("Nest 6", "styleListItem6"))
      .addSeparator()            
      .addItem('Bold', 'styleBold')
      .addItem('Italic', 'styleItalic')
      .addItem('Underline', 'styleUnderline')
      .addSeparator()
      .addItem('Image', 'styleImage')
      .addItem('Link', 'styleLink')
      .addSeparator()
      .addItem('Comment', 'styleComment')
      .addSeparator()
      .addItem('Clean', 'cleanStyle')
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
