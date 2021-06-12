function myFunction() {
    var generalCalendarID = 'urspadvisoryboard@gmail.com';
    var classRepCalendarID = 'fe6ah9vk3lpupi9gedhpqqg51g@group.calendar.google.com';
    var mentorCalendarID = 'qleoc35j6fsifab4kdbhq8hr7s@group.calendar.google.com';
    var currentCalendarID;
    var eventCal;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];

    // This represents ALL the data
    var range = sheet.getDataRange();
    var tasks = range.getValues();

    for (x=0; x<tasks.length;x++)
    {
        var duty = tasks[x];
        dutyNumber = duty[1];
        if (dutyNumber.substring(0,1) == 1)
          currentCalendarID = classRepCalendarID;
        else if (dutyNumber.substring(0,1) == 3)
          currentCalendarID = mentorCalendarID;
        else
          currentCalendarID = generalCalendarID;
        var dutyName = duty[2];
        var startDate = new Date(duty[3]);
        var endDate = new Date(duty[4]);
        eventCal = CalendarApp.getCalendarById(currentCalendarID);
        eventCal.createEvent(dutyName, startDate, endDate);
    }
}
