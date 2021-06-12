function onOpen() {
    SpreadsheetApp.getUi().createMenu('Info Sheet')
        .addItem('Attach sheets', 'attachSheets_')
        .addToUi();
}

function attachSheets_() {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName('Sheet1');
    var ss2 = SpreadsheetApp.openById("1S7ioMUlxnEoNNHYbkW_LA2TwVip0HfKHMmEs9vDkDkk");
    var sheet2 = ss2.getSheetByName('Practice');
    var mentorvalues = sheet.getDataRange().getValues();
    var menteevalues = sheet2.getDataRange().getValues();

    var headers = mentorvalues[0];
    var mentee1name = headers.indexOf("Mentee #1 Name");
    var mentee1PDF = headers.indexOf("Mentee #1 PDF");
    var mentee2name = headers.indexOf("Mentee #2 Name");
    var mentee2PDF = headers.indexOf("Mentee #2 PDF");
    var mentee3name = headers.indexOf("Mentee #3 Name");
    var mentee3PDF = headers.indexOf("Mentee #3 PDF");
  
    var mentee1URL = "";
    var mentee2URL = "";
    var mentee3URL = "";
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var values = sheet.getDataRange().getValues();
  
    var attachmentIndex = headers.indexOf("Attachment");

    for (var k = 1; k < menteevalues.length; k++) {
        var rowDatamentee = menteevalues[k];
        var attachment = "";
        for (var i = 1; i < mentorvalues.length; i++) {
            var rowDatamentor = mentorvalues[i];
            var menteeIndex = rowDatamentee.length - 1;
            if (rowDatamentee[0] == rowDatamentor[mentee1name]) {
                sheet.getRange(i + 1, mentee1PDF + 1).setValue(rowDatamentee[menteeIndex]);
                mentee1URL = rowDatamentee[menteeIndex];
                //console.log(rowDatamentee[0]);
            } else if (rowDatamentee[0] == rowDatamentor[mentee2name]) {
                sheet.getRange(i + 1, mentee2PDF + 1).setValue(rowDatamentee[menteeIndex]);
                mentee2URL = rowDatamentee[menteeIndex];
                //console.log(rowDatamentee[0]);
            } else if (rowDatamentee[0] == rowDatamentor[mentee3name]) {
                sheet.getRange(i + 1, mentee3PDF + 1).setValue(rowDatamentee[menteeIndex]);
                mentee3URL = rowDatamentee[menteeIndex];
                //console.log(rowDatamentee[0]);
            }
          //attachment = mentee1URL + ", " + mentee2URL + ", " + mentee3URL;
      //sheet.getRange(i, attachmentIndex + 1).setValue(attachment);
        }
    }
}