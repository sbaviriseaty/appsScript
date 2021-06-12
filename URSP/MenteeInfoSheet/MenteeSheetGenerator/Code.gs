function onOpen() {
    SpreadsheetApp.getUi().createMenu('Info Sheet')
        .addItem('Make sheets', 'start')
        .addItem('Make sheets (version two)', 'complete2')
        .addToUi();
}

var DocsFolder = DriveApp.getFolderById("13E0rDaKP9SPxqUXgsFO8F9sw_Edk3WgB"); 
var PDFFolder = DriveApp.getFolderById("1tUg62i8a3tdPPXX1KUxZtetumNA1jiXP");

function complete2() {
 
  // Load the Google Doc template file
  var templateId = "13iCKWxts89bIJZnQtV3p7Pz5K9CzFgYzmTQ_ZWhLh-0";
  var template = DriveApp.getFileById(templateId);
  
  // Get all mentee data from the spreadsheet and identify the headers
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var NameIndex = headers.indexOf("Name");
  var MajorIndex = headers.indexOf("Major");
  var HometownIndex = headers.indexOf("Hometown");
  var AnimalIndex = headers.indexOf("Favorite Animal");
  var ResearchIndex = headers.indexOf("Research Interests");
  var LinkIndex = headers.indexOf("URL");
    
  // Iterate through each row to capture individual details
  for (var i = 1; i < values.length; i++) {
    var rowData = values[i];
    var Name = rowData[NameIndex];
    var Major = rowData[MajorIndex];
    var Hometown = rowData[HometownIndex];
    var Animal = rowData[AnimalIndex];
    var Research = rowData[ResearchIndex];
    
    // Make a copy of the sheet template and rename it with mentee name
    var sheetId = template.makeCopy(DocsFolder).setName(Name + " Info Sheet").getId();      
    var doc = DocumentApp.openById(sheetId);
    var body = doc.getBody();    
    
    // Replace placeholder values with actual mentee related details
    body.replaceText("{{Name}}", Name);
    body.replaceText("{{Major}}", Major);
    body.replaceText("{{Hometown}}", Hometown);
    body.replaceText("{{Favorite Animal}}", Animal);
    body.replaceText("{{Research Interests}}", Research);
    
    // Update the spreadsheet with the new sheet Id and status
    doc.saveAndClose();
    docblob = doc.getAs('application/pdf');
    docblob.setName(doc.getName() + ".pdf");
    var PDF = DriveApp.createFile(docblob);
    PDF.moveTo(PDFFolder);
    var doc = DriveApp.getFileById(doc.getId());
    doc.moveTo(DocsFolder);
    url = "https://drive.google.com/file/d/" + PDF.getId() + "/view?usp=sharing";
    sheet.getRange(i + 1, LinkIndex + 1).setValue(url);
    SpreadsheetApp.flush();
  }
}

function start() {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName('Practice');
    var values = sheet.getDataRange().getValues();

    var headers = values[0];
    var LinkIndex = headers.indexOf("URL");

    // Iterate through each row/student to capture individual details
    for (var i = 1; i < values.length; i++) {
        var rowData = values[i];
        url = makeDoc(sheet, rowData);
        sheet.getRange(i + 1, LinkIndex + 1).setValue(url);
        SpreadsheetApp.flush();
    }
}

function makeDoc(sheet, rowData) {
    var doc = DocumentApp.create('Mentee Info Sheet for ' + rowData[0]);
    var body = doc.getBody();
    var headings = sheet.getDataRange()
        .offset(0, 0, 1)
        .getValues()[0];
    console.log(headings);
    for (var i = 0; i < rowData.length - 1; i++) {
        body.appendParagraph(headings[i] + ": " + rowData[i]);
    }
    body.insertParagraph(0, doc.getName())
        .setHeading(DocumentApp.ParagraphHeading.HEADING3);
    doc.saveAndClose();
    docblob = doc.getAs('application/pdf');
    docblob.setName(doc.getName() + ".pdf");
    var PDF = DriveApp.createFile(docblob);
    PDF.moveTo(PDFFolder);
    var doc = DriveApp.getFileById(doc.getId());
    doc.moveTo(DocsFolder);
    url = "https://drive.google.com/file/d/" + PDF.getId() + "/view?usp=sharing";
    return url;
}