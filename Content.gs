//Style Stuff
var space = " ";
var lineBreak = "\n";


// Symbol Checkers

var listItemChecker = new Array(/^@{1}\s/, /^@{2}\s/, /^@{3}\s/, /^@{4}\s/, /^@{5}\s/, /^@{6}\s/); //Headers from 1-6
var headerChecker = new Array(/^\${1}\s/, /^\${2}\s/, /^\${3}\s/, /^\${4}\s/, /^\${5}\s/, /^\${6}\s/); //List items from 1-6
var urlChecker = /w{3}|http/; //Searches for either www or http
var commentChecker = /\/{2}\s/;

var boldChecker = /\[b\]\S.+?\[\/b\]/gm;
var italicChecker = /\[i\]\S.+?\[\/i\]/gm;
var underlineChecker = /\[u\]\S.+?\[\/u\]/gm

var docTitlePlaceholder = /![xX]!/; //X with any lower or upper caps
var productMatrixPlaceholder = /![pP][rR][oO][dD][uU][cC][tT][mM][aA][tT][rR][iI][xX]!/; //ProductMatrix with any lower or upper caps
var productListPlaceholder = /![pP][rR][oO][dD][uU][cC][tT][lL][iI][sS][tT]!/; //ProductMatrix with any lower or upper caps
var linePlaceholder = /![lL][iI][nN][eE]!/;

//Content generation methods

function addElement(element) {
  
  if(element != "") {
    
    var newElement;
    
    if(element.search(listItemChecker[0]) >= 0) {
      element = removeSymbol(element, listItemChecker[0]); newElement = body.appendListItem(element).setGlyphType(DocumentApp.GlyphType.BULLET).setNestingLevel(0); }
    else if(element.search(listItemChecker[1]) >= 0) {
      element = removeSymbol(element, listItemChecker[1]); newElement = body.appendListItem(element).setGlyphType(DocumentApp.GlyphType.HOLLOW_BULLET).setNestingLevel(1); }
    else if(element.search(listItemChecker[2]) >= 0) {
      element = removeSymbol(element, listItemChecker[2]); newElement = body.appendListItem(element).setGlyphType(DocumentApp.GlyphType.HOLLOW_BULLET).setNestingLevel(2); }
    else if(element.search(listItemChecker[3]) >= 0) {
      element = removeSymbol(element, listItemChecker[3]); newElement = body.appendListItem(element).setGlyphType(DocumentApp.GlyphType.HOLLOW_BULLET).setNestingLevel(3); }
    else if(element.search(listItemChecker[4]) >= 0) {
      element = removeSymbol(element, listItemChecker[4]); newElement = body.appendListItem(element).setGlyphType(DocumentApp.GlyphType.HOLLOW_BULLET).setNestingLevel(4); }
    else if(element.search(listItemChecker[5]) >= 0) {
      element = removeSymbol(element, listItemChecker[5]); newElement = body.appendListItem(element).setGlyphType(DocumentApp.GlyphType.HOLLOW_BULLET).setNestingLevel(5); }
    else if(element.search(headerChecker[0]) >= 0) {
      element = removeSymbol(element, headerChecker[0]); newElement = body.appendParagraph(element).setHeading(DocumentApp.ParagraphHeading.HEADING1); }
    else if(element.search(headerChecker[1]) >= 0) {
      element = removeSymbol(element, headerChecker[1]); newElement = body.appendParagraph(element).setHeading(DocumentApp.ParagraphHeading.HEADING2); }
    else if(element.search(headerChecker[2]) >= 0) {
      element = removeSymbol(element, headerChecker[2]); newElement = body.appendParagraph(element).setHeading(DocumentApp.ParagraphHeading.HEADING3); }
    else if(element.search(headerChecker[3]) >= 0) {
      element = removeSymbol(element, headerChecker[3]); newElement = body.appendParagraph(element).setHeading(DocumentApp.ParagraphHeading.HEADING4); }
    else if(element.search(headerChecker[4]) >= 0) {
      element = removeSymbol(element, headerChecker[4]); newElement = body.appendParagraph(element).setHeading(DocumentApp.ParagraphHeading.HEADING5); }
    else if(element.search(headerChecker[5]) >= 0) {
      element = removeSymbol(element, headerChecker[5]); newElement = body.appendParagraph(element).setHeading(DocumentApp.ParagraphHeading.HEADING6); }
    else if(element.search(commentChecker) >= 0) {
      //Do nothing for comments 
    }
    else {
      newElement = body.appendParagraph(element).setHeading(DocumentApp.ParagraphHeading.NORMAL); }
    
    if(element.search(urlChecker) >=0 && element.search(commentChecker) < 0)
      newElement.setLinkUrl(element);
      
  }
}


function sortData() {
  
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  
  var activeSheet = sheet.getActiveSheet();
  
  for(i=1; i<=(lastColumn-mainAreaStart+1); i++) {
    
    mainAreaList[i - 1] = activeSheet.getRange(1, (i + mainAreaStart -1), lastRow, (lastColumn - mainAreaStart +1)).getValues();
    
  }
  
  for(i=0;i<mainAreaList.length;i++)
    for(j=0;j<mainAreaList[i].length;j++)
      if(mainAreaList[i][j] != "")
        documentList.push(mainAreaList[i][j]);
  
  
  return documentList[0];
}

function removeSymbol(text, symbol) {
  
  var newText = text.split(symbol);
  
  newText = newText[1];
  
  return newText;
}


function createProductList(childIndex) {
  
  body.removeChild(body.getChild(childIndex));
  
  productList = sheet.getRange(productListRange[0] + ":" + productListRange[1]).getValues();
  
  for(i=0;i<productList.length;i++) {
    
    var productLabel;
    
    if(productList[i][0] != "")
      productLabel = productList[i][0] + ": ";
    else
      productLabel = "";
    
    if(productList[i][1] != "")
      body.insertListItem(childIndex + i, productLabel + productList[i][1]).setLinkUrl(productList[i][2]);
  }
  
}

function createProductMatrix(childIndex) {
  
  body.removeChild(body.getChild(childIndex));
  
  productList = sheet.getRange(productListRange[0] + ":" + productListRange[1]).getValues();
  
  body.insertParagraph(childIndex, ""); // First space
  var productTable = body.insertTable(childIndex);
  
  var productHeader = productTable.appendTableRow();
  var productOverview = productTable.appendTableRow();
  var productPros = productTable.appendTableRow();
  var productCons = productTable.appendTableRow();
  
  productHeader.appendTableCell("Products");
  productOverview.appendTableCell("Overview");
  productPros.appendTableCell("Pros");
  productCons.appendTableCell("Cons");
  
  for(i=0;i<productList.length;i++)
    if(productList[i][1] != "") {
      productHeader.appendTableCell(productList[i][1]).setLinkUrl(productList[i][2]);
      productOverview.appendTableCell("");
      productPros.appendTableCell("");
      productCons.appendTableCell("");
    }
  
  body.insertParagraph(childIndex, ""); // Last space
  
}

function formatURL() {
  
  for(counter = 4; counter <=49; counter ++) {
    
    var urlCellFirst = sheet.getRange("B" + counter).getValue();
    
    if (urlCellFirst.indexOf("amazon") >= 0) {
      urlCellFirst = urlCellFirst.split("/ref");
      urlCellFirst = urlCellFirst[0] + fgTag;
      sheet.getRange("C" + counter).setValue(urlCellFirst);
    }
    else if (urlCellFirst.indexOf("walmart") >= 0) {
      urlCellFirst = urlCellFirst.split("?");
      sheet.getRange("C" + counter).setValue(urlCellFirst[0]);
    }  
  }
  
}


function checkIntext(textPos) {
  
  var bodyText = body.getChild(textPos).asText();
  
  var results = new Array(); 
  
  results = bodyText.getText().match(boldChecker);
  formatIntext(bodyText, results, "bold");
  
  results = bodyText.getText().match(italicChecker);
  formatIntext(bodyText, results, "italic");
  
  results = bodyText.getText().match(underlineChecker);
  formatIntext(bodyText, results, "underline");
  
}


function formatIntext(bodyText, results, style) {
  
  if(results != null) {
   
    for (i=0;i<results.length;i++) {
      
      var start = bodyText.getText().indexOf(results[i]);
      var end = start + results[i].length - 1;
      
      var tag1Start = start;
      var tag1End = start + 2;
      
      var tag2Start = end - 3 - 3;
      var tag2End = end - 3;
      
      if(style == "bold")
        bodyText.setBold(start, end, true);
      else if(style == "italic")
        bodyText.setItalic(start, end, true);
      else if(style == "underline")
        bodyText.setUnderline(start, end, true);
      
      bodyText.deleteText(tag1Start, tag1End);
      bodyText.deleteText(tag2Start, tag2End);
    }
  }
}
