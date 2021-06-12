function generateAwardsPresentation() {
  var dataSpreadsheetUrl = "https://docs.google.com/spreadsheets/d/1hCXchELKo4SKet3w_ElHxko2qvEiJHRiJbkMlS-CJNU/edit"; //make sure this includes the '/edit at the end
  var ss = SpreadsheetApp.openByUrl(dataSpreadsheetUrl);
  var deck = SlidesApp.getActivePresentation();
  var sheet = ss.getSheetByName('modified-results');
  var values = sheet.getRange('A2:E46').getValues();
  var slides = deck.getSlides();
  var templateSlide = slides[1];
  var presLength = slides.length;
  
  values.forEach(function(page){
  if(page[0]){
    
   var eventName = page[0];
   var eventDivision = page[1];
   var Team1Name = page[2];
   var Team2Name = page[3];
   var Team3Name = page[4];  
    
   templateSlide.duplicate(); //duplicate the template page
   slides = deck.getSlides(); //update the slides array for indexes and length
   newSlide = slides[2]; // declare the new page to update
    
    
   var shapes = (newSlide.getShapes());
     shapes.forEach(function(shape){
       shape.getText().replaceAllText('{{event name}}',eventName);
       shape.getText().replaceAllText('{{event division}}',eventDivision);
       shape.getText().replaceAllText('{{team 1 name}}',Team1Name);
       shape.getText().replaceAllText('{{team 2 name}}',Team2Name);
       shape.getText().replaceAllText('{{team 3 name}}',Team3Name);
    }); 
   presLength = slides.length; 
   newSlide.move(presLength); 
  } // end our conditional statement
  }); //close our loop of values

//Remove the template slide
templateSlide.remove();
  
}
