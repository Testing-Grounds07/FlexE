//Style Stuff
var space = " ";
var lineBreak = "\n";


// Symbol Checkers

var listItemChecker = new Array(/^@{1}\s/, /^@{2}\s/, /^@{3}\s/, /^@{4}\s/, /^@{5}\s/, /^@{6}\s/); //Headers from 1-6
var listItemAllChecker = /^@+?\s/gm;
var headerChecker = new Array(/^\${1}\s/, /^\${2}\s/, /^\${3}\s/, /^\${4}\s/, /^\${5}\s/, /^\${6}\s/, /^!\${1}\s/); //List items from 1-6
var headerAllChecker = /^\$+?\s/gm;
var urlChecker = /w{3}|http/; //Searches for either www or http
var commentChecker = /\/{2}\s/;
var tagChecker = /(@|\$|\!\$|\/)+?\s/gm;

var boldChecker = /\[b\]\S.+?\[\/b\]/gm;
var italicChecker = /\[i\]\S.+?\[\/i\]/gm;
var underlineChecker = /\[u\]\S.+?\[\/u\]/gm;
var linkChecker = /\[l=\S+?\]\S.+?\[\/l\]/gm;
var imgChecker = /\[img\]\S+?\[\/img\]/gm;
var bracketChecker = new Array(/\[\S+?\]/gm, /\[\/\S+?\]/gm);

var formatingChecker = /\#\S.+?\#/gm;

var linkURL = /\[l=\S.+?\]/gm;

var docTitlePlaceholder = /![xX]!/; //X with any lower or upper caps
var productMatrixPlaceholder = /![pP][rR][oO][dD][uU][cC][tT][mM][aA][tT][rR][iI][xX]!/; //ProductMatrix with any lower or upper caps
var productListPlaceholder = /![pP][rR][oO][dD][uU][cC][tT][lL][iI][sS][tT]!/; //ProductMatrix with any lower or upper caps
var linePlaceholder = /![lL][iI][nN][eE]!/;
var breakPlaceholder = /![bB][rR][eE][aA][kK]!/;

// Formating Tags
var listItemTag = [["@ ",""], ["@@ ",""], ["@@@ ",""], ["@@@@ ",""], ["@@@@@ ",""], ["@@@@@@ ",""]];

var headerTag  = [["$ ",""], ["$$ ",""], ["$$$ ",""], ["$$$$ ",""], ["$$$$$ ",""], ["$$$$$$ ",""], ["!$ ",""]];

var anyStyleTags = /(@+?|\$+?|\!\$|\/\/)\s/gm;

var commentTag = ["// ",""];
var boldTags  = ["[b]","[/b]"];
var italicTags = ["[i]","[/i]"];
var underlineTags = ["[u]","[/u]"];
var imgTags = ["[img]","[/img]"];
var linkTags = ["[l=]","[/l]"];

var anyBracketTags = new Array(/\[\S+?\]/, /\[\/\S+?\]/);

var formatTag = /\#/gm;


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
    else if(element.search(headerChecker[6]) >= 0) {
      element = removeSymbol(element, headerChecker[6]); newElement = body.appendParagraph(element).setHeading(DocumentApp.ParagraphHeading.TITLE); }
    else if(element.search(commentChecker) >= 0) {
      //Do nothing for comments 
    }
    else {
      newElement = body.appendParagraph(element).setHeading(DocumentApp.ParagraphHeading.NORMAL); }
    
    if(element.search(urlChecker) >=0 && element.search(commentChecker) < 0 && element.search(linkChecker) < 0 && element.search(imgChecker) < 0)
      newElement.setLinkUrl(element);
      
  }
}

function removeCell(cellTarget) {
  
  sheet.getRange(cellTarget).setValue("");
  
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
  
  var bodyText = body.getChild(textPos);
  
  var results = new Array(); 
  var test = bodyText.getText();
  results = bodyText.getText().match(boldChecker);
  formatIntext(bodyText, results, "bold");
  
  results = bodyText.getText().match(italicChecker);
  formatIntext(bodyText, results, "italic");
  
  results = bodyText.getText().match(underlineChecker);
  formatIntext(bodyText, results, "underline");
  
  results = bodyText.getText().match(linkChecker);
  formatIntext(bodyText, results, "link");
  
  results = bodyText.getText().match(imgChecker);
  formatIntext(bodyText, results, "image");
  
  results = bodyText.getText().match(formatingChecker);
  formatIntext(bodyText, results, "format");
  
}


function formatIntext(bodyText, results, style) {
  
  if(results != null) {
   
    for (i=0;i<results.length;i++) {
      
      var start = bodyText.getText().indexOf(results[i]);
      var end = start + results[i].length - 1;
      
      var tag1Start = start;
      var tag1End = start + 2;
      
      var tag2Start = end - 3;
      var tag2End = end ;
      
      if(style == "bold")
        bodyText.asText().setBold(start, end, true);
      else if(style == "italic")
        bodyText.asText().setItalic(start, end, true);
      else if(style == "underline")
        bodyText.asText().setUnderline(start, end, true);
      else if(style == "link") {
       
        var linkString = bodyText.asText().getText().match(linkURL);
        linkString = linkString[0].substr(3);
        linkString = linkString.slice(0, -1);
        
        bodyText.asText().setLinkUrl(start, end, linkString);
        
        tag1End = tag1End + linkString.length + 1;
        
      }
      else if(style == "image") {
        
        var imgURL = bodyText.asText().getText().match(imgChecker);
        
        imgURL = imgURL[0].substr(5);
        imgURL = imgURL.slice(0, -6);
        
        var resp = UrlFetchApp.fetch(imgURL);
        var image = resp.getBlob();
        
        bodyText.clear();
        bodyText.appendInlineImage(image);
        
        tag1End += 2;
        tag2Start -= 2;
        
      }
      else if(style == "format") {
        
        tag1End -= 2;
        tag2Start += 3;
        
      }
      
      if(style != "image") {
        bodyText.asText().deleteText(tag2Start, tag2End);
        bodyText.asText().deleteText(tag1Start, tag1End);
      }
    }
  }
}

function reformatTag(targetTag, contentTags) {

  if(contentTags.length > 0) {
    if(targetTag.search(formatingChecker) < 0) {
      if(Array.isArray(targetTag.match(anyStyleTags))) {
        
        var headTag = targetTag.match(anyStyleTags);
        targetTag = targetTag.replace(headTag[0], "");
        targetTag = headTag[0] + contentTags[0] + targetTag + contentTags[1];
        
      }
      else
        targetTag = contentTags[0] + targetTag + contentTags[1];
    }
    else {
      
      var portions = targetTag.match(formatingChecker);
      
      for(l=0; l<portions.length; l++) {
        var oldPortion = portions[l]; 
        portions[l] = portions[l].split(formatTag);
        
        portions[l][0] = contentTags[0];
        portions[l][2] = contentTags[1];
        
        portions[l] = portions[l][0] + portions[l][1] + portions[l][2];
        
        targetTag = targetTag.replace(oldPortion, portions[l]);
      }
    }
  }
  else {
    
    targetTag = targetTag.replace(tagChecker, "");
    targetTag = targetTag.replace(bracketChecker[0], "");
    targetTag = targetTag.replace(bracketChecker[1], "");
    targetTag = targetTag.replace(formatTag, "");
    
  }
  
  return targetTag;
  
}

function addTags(contentTags) {
  
  var sel = SpreadsheetApp.getActive().getSelection().getActiveRangeList().getRanges();
  
  for(i=0; i < sel.length; i++) {
    
    var cellRange = sheet.getRange(sel[i].getA1Notation());
    var rangeValues = cellRange.getValues();
    
    for(j=0; j < rangeValues.length; j++) {
      
      for(k=0; k < rangeValues[j].length; k++) {
        var value = rangeValues[j][k].toString();
        
        rangeValues[j][k] = reformatTag(value, contentTags);
        
      }
      
    }
    cellRange.setValues(rangeValues);

  }
  
}

function scrubTags() {
  
  
  var sel = SpreadsheetApp.getActive().getSelection().getActiveRangeList().getRanges();
  
  for(i=0; i < sel.length; i++) {
    
    var cellRange = sheet.getRange(sel[i].getA1Notation());
    var rangeValues = cellRange.getValues();
    
    for(j=0; j < rangeValues.length; j++) {
      
      for(k=0; k < rangeValues[j].length; k++) {
        var value = rangeValues[j][k].toString();
        
        rangeValues[j][k] = reformatTag(value, []); // Sends empty array to remove tags.
        
      }
      
    }
    
    cellRange.setValues(rangeValues);
  }
  
}

// Content Style Commands

function styleListItem1() {
  addTags(listItemTag[0]);
}

function styleListItem2() {
  addTags(listItemTag[1]);
}

function styleListItem3() {
  addTags(listItemTag[2]);
}

function styleListItem4() {
  addTags(listItemTag[3]);
}

function styleListItem5() {
  addTags(listItemTag[4]);
}

function styleListItem6() {
  addTags(listItemTag[5]);
}


function styleHeader1() {
  addTags(headerTag[0]);
}

function styleHeader2() {
  addTags(headerTag[1]);
}

function styleHeader3() {
  addTags(headerTag[2]);
}

function styleHeader4() {
  addTags(headerTag[3]);
}

function styleHeader5() {
  addTags(headerTag[4]);
}

function styleHeader6() {
  addTags(headerTag[5]);
}

function styleHeaderTitle() {
  addTags(headerTag[6]);
}


function styleComment() {
  addTags(commentTag);
}

function styleBold() {
  addTags(boldTags);
}

function styleItalic() {
  addTags(italicTags);
}

function styleUnderline() {
  addTags(underlineTags);
}

function styleImage() {
  addTags(imgTags);
}

function styleLink() {
  addTags(linkTags);
}


function cleanStyle() {
  scrubTags();
}

function tempUpdate() {
  
  updateDocument();
  
}
