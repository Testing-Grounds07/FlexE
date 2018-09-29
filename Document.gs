var sheet = SpreadsheetApp.getActiveSpreadsheet();
var data = sheet.getDataRange().getValues();

//Info Cells
var docIDCell = "B2";
var docLinkCell = "C2";

//Document Info
var docTitle = sheet.getRange("B1").getValue();
var docType = sheet.getRange("C1").getValue();

var doc;
var docID = sheet.getRange(docIDCell).getValue();
var body;
var lists;


//Template/Overview/archiveInfo Information
var templateSheetPosition = 1; // The location of the template sheet at the bottom of the spreadsheet
var templateInfoLocations = new Array("B1", "C1");
var templateSheet = sheet.getSheets()[templateSheetPosition];

var templateTitle = templateSheet.getRange(templateInfoLocations[0]).getValue();
var templateType = templateSheet.getRange(templateInfoLocations[1]).getValue();

var templateName = templateType + " - " + templateTitle;
var overviewName = "Overview";

var archiveID = sheet.getSheets()[0].getRange(docIDCell).getValue(); //ID for archive spreadsheet.
var archiveSheet = SpreadsheetApp.openById(archiveID);

var fgTag = "/&tag=fitgo-20";


//Data Lists & Ranges

var mainAreaStart = 4;

var mainAreaList = new Array();
var documentList = new Array();

var productListRange = new Array("A4","C");
var productList = new Array();

var keywordsLink;


//Retrive existing document information

if(docID != "" && isStaticSheet("both") == false) {
  doc = DocumentApp.openById(docID);
  body = doc.getBody();

}


//Document methods

function createNewSpreadsheet() {
  
  var newSheet = sheet.setActiveSheet(sheet.getSheetByName(templateName)).copyTo(sheet);
  sheet.setActiveSheet(newSheet);
    
}

function newDocument() {
  
  doc = DocumentApp.create("Untitled");
  body = doc.getBody();
  docID = doc.getId();
  docPermission = DriveApp.getFileById(docID);
  docPermission.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.COMMENT);
  sheet.getRange(docIDCell).setValue(docID);
  
  //Create Open link
  
  var openLink = '=HYPERLINK("https://docs.google.com/document/d/' + docID + '/","Open Document")';
  sheet.getRange(docLinkCell).setValue(openLink);
  
}

function updateDocument() {
  
  renameSheet(docType + " - " + docTitle);
    
  if(docID == "" && isStaticSheet("both") == false) {
    
    newDocument();

  }
  
  if(isStaticSheet("overview") == false)
    renameDocument(docType + " - " + docTitle);
  
  if(isStaticSheet("both") == false) {
    clearDocument();
    
    sortData(); //Upload spreadsheet content
    
    body.appendParagraph(docType + " - " + docTitle).setHeading(DocumentApp.ParagraphHeading.TITLE);
    
    body.removeChild(body.getChild(0)); //Remove extra space at top
    
    for(i=0;i<documentList.length;i++) //Add dynamic content
      addElement(documentList[i][0]);
    
    
    //List Item Styling && ToC
    for(z=0;z < body.getNumChildren();z++) {
      
      reviewElement = body.getChild(z);
      if(reviewElement.getType() == DocumentApp.ElementType.LIST_ITEM) {
        var reviewList = reviewElement.asListItem();
        
        if(reviewList.getNestingLevel() == 0)
          reviewList.setGlyphType(DocumentApp.GlyphType.BULLET);
        else
          reviewList.setGlyphType(DocumentApp.GlyphType.HOLLOW_BULLET);
        
        checkIntext(z);
      }
      else if(reviewElement.getType() == DocumentApp.ElementType.PARAGRAPH) {
        var reviewItem = reviewElement.asParagraph();
        
        if(reviewItem.getText().search(docTitlePlaceholder) >= 0)
          reviewElement.asParagraph().editAsText().replaceText(reviewItem.getText().match(docTitlePlaceholder)[0], docTitle.toLowerCase());
        else if(reviewItem.getText().search(productMatrixPlaceholder) >= 0)
          createProductMatrix(body.getChildIndex(reviewElement));
        else if(reviewItem.getText().search(productListPlaceholder) >= 0)
          createProductList(body.getChildIndex(reviewElement));
        
        checkIntext(z);
      }
      
      
    }
  }
  
}

function clearDocument() {
  
  if(docID != "") {
    body.appendParagraph('');
  
    body.clear();
  }
  
}

function renameDocument(documentTitle) {
  
  if(docID != "")
    doc.setName(documentTitle);
  
}

function renameSheet(documentTitle) {
  
  sheet.getActiveSheet().setName(documentTitle);
  
}

function remove(removeTarget) {
  
  if(removeTarget == "document") {
    if(docID != "") {
      var file = DriveApp.getFileById(docID);
      file.setTrashed(true);
      
      sheet.getRange(docIDCell).setValue("");
      sheet.getRange(docLinkCell).setValue("");
    }
  }
  else if(removeTarget == "spreadsheet") {
  
  if(isStaticSheet("both") == false) //Protect template/overview sheets
    sheet.deleteActiveSheet();
  }
}


function archive() {
  
  if(sheet.getActiveSheet().getName() != templateName || sheet.getActiveSheet().getName() != overviewName) //Protect template/overview sheets
    sheet.getActiveSheet().copyTo(archiveSheet).setName(docType + " - " + docTitle + " (Archived)"); remove("spreadsheet");
  
}


function isStaticSheet(checkerType) {
  
  if(checkerType == "both") {
    templateTitle = templateSheet.getRange(templateInfoLocations[0]).getValue();
    templateType = templateSheet.getRange(templateInfoLocations[1]).getValue();

    templateName = templateType + " - " + templateTitle;
    if(sheet.getActiveSheet().getName() != templateName && sheet.getActiveSheet().getName() != overviewName) {
      return false;
    }
    else {
      return true;
    }
  }
  else if(checkerType == "overview") {
    if(sheet.getActiveSheet().getName() != overviewName) {
      return false;
    }
    else {
      return true;
    }
  }
  
}
