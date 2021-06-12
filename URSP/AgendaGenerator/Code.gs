function onOpen() {
  var ui = SlidesApp.getUi();
  ui.createMenu('Create Agenda')
      .addItem('Pull Data', 'generateURSPGBMAgenda')
      .addToUi();
}

function generateURSPGBMAgenda() {
  var dataSpreadsheetUrl = "https://docs.google.com/spreadsheets/d/1-bGvkObZPXY9PVQIKtDyCMZoJmbxKS-iK4noGHNPm10/edit"; //make sure this includes the '/edit at the end
  var ss = SpreadsheetApp.openByUrl(dataSpreadsheetUrl);
  var deck = SlidesApp.getActivePresentation();
  var sheet = ss.getSheetByName('april'); //CHANGE THIS NAME FOR EACH NEW PRESENTATION
  var dates = sheet.getRange('A1:A2').getValues();
  var values = sheet.getRange('A5:E11').getValues();
  var slides = deck.getSlides();
  var templateSlide = slides[1];
  var presLength = slides.length;

  var month = dates[0];
  var date = dates[1];

  var titleSlide = slides[0];
  var shapes = (titleSlide.getShapes());
     shapes.forEach(function(shape){
       shape.getText().replaceAllText('{{Month}}',month);
       shape.getText().replaceAllText('{{Date}}',date);
    }); 
  
  values.forEach(function(page){
  if(page[0]){
    
   var committee = page[0];
   var update1 = page[1];
   var update2 = page[2];
   var update3 = page[3];
   var update4 = page[4];
      
   templateSlide.duplicate(); //duplicate the template page
   slides = deck.getSlides(); //update the slides array for indexes and length
   newSlide = slides[2]; // declare the new page to update
   
       
   var shapes = (newSlide.getShapes());
     shapes.forEach(function(shape){
       shape.getText().replaceAllText('{{Committee}}',committee);
       shape.getText().replaceAllText('{{Update1}}',update1);
       shape.getText().replaceAllText('{{Update2}}',update2);
       shape.getText().replaceAllText('{{Update3}}',update3);
       shape.getText().replaceAllText('{{Update4}}',update4);
    }); 
   presLength = slides.length; 
   newSlide.move(presLength); 
  } // end our conditional statement
  }); //close our loop of values

//Remove the template slide
templateSlide.remove();
  
}
