/**
 * This is triggered whenever there is a change to the calendar
 * When there is a change, it searches for events that include "#agenda"
 * in the decrisption.
 *
 */
function onCalendarChange() {
  // Get recent events with the tag
  const now = new Date();
  const events = CalendarApp.getEvents(
      now,
      new Date(now.getTime() + 2 * 60 * 60 * 1000000),
      {search: '#agenda'},
  );

  const folderId = '1ZWbuWGXAfe4Fccdql4EfKlrP0G0BrDp-'; //2021 Meetings Folder
  const templateId = '1eiHIgOAlgyNzxd0PXmneVG1FbAHv-Wc3vjKxLDRu55A'; //General Meetings Minutes Template

  const folder = DriveApp.getFolderById(folderId);

  // Go through any events found
  for (i = 0; i < events.length; i++) {
    const event = events[i];

    // Confirm whether the event has the tag
    let description = event.getDescription();
    if (description.search('#agenda') == -1) continue;

    // Only work with events created by the owner of this calendar
    if (event.isOwnedByMe()) {
      if (DriveApp.getFilesByName('Agenda for ' + event.getTitle()) == 1)
        continue;
      // Create a new document from the template, for an agenda for this event
      const newDoc = DriveApp.getFileById(templateId).makeCopy();
      newDoc.setName('Agenda for ' + event.getTitle());

      const file = DriveApp.getFileById(newDoc.getId());
      folder.addFile(file);

      // Replace the tag with a link to the agenda document
      const agendaUrl = 'https://docs.google.com/presentation/d/' + newDoc.getId();
      description = description.replace(
          '#agenda',
          '<a href=' + agendaUrl + '>Agenda for ' + event.getTitle() + '</a>',
      );
      event.setDescription(description);      
    }
  }
  return;
}

/**
 * Register a trigger to watch for calendar changes.
 */
function setUp() {
  var email = Session.getActiveUser().getEmail();
  ScriptApp.newTrigger("onCalendarChange").forUserCalendar(email).onEventUpdated().create();
}
