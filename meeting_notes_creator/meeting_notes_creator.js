/*
Date: January 2020
Author: Marquin Smith

Description: Program to create a google doc for every meeting for today.
*/

// Helper functions

function getEventsToday() {
    // collect all calendar events for today.
    var start = new Date();
    start.setHours(0,0,0,0);
    
    var end = new Date();
    end.setHours(23,59,59,999);

    var events = CalendarApp.getDefaultCalendar().getEvents(start, end);

    return events;
    
  }


function getEventDetails(some_event) {
    // given a calendar event will return the 
    // following in an array
    // start time
    // event title
    // event description
    // attendees
    var my_desc = some_event.getDescription();
    console.log(my_desc);

    var my_start = some_event.getStartTime();
    //console.log(my_start);
    //console.log(my_start.getFullYear());
    //console.log(my_start.getDate());
    

    var my_title = some_event.getTitle();
    //console.log(my_title);

    var my_attendees = some_event.getGuestList(true);
    //console.log(my_attendees);

    var toReturn = [my_start,
                    my_title,
                    my_desc,
                    my_attendees]

    return toReturn
    
  }

function prependZeroIfNecessary(a_number) {
    // converts a number to string
    // and appends a zero to the beginning if the length is less than 2
    a_number = a_number.toString();
    if (a_number.length < 2) {
        var a_number = "0" + a_number
    }

    return a_number
}

function parseDatesToFormat(start_date) {
    // given a date will return a string 
    // to form part of the document title in the form
    // yyyymmdd - hhmm - 

    var my_year = start_date.getFullYear();
    my_year = my_year.toString();

    var my_month = start_date.getMonth() + 1;
    my_month = prependZeroIfNecessary(my_month);

    var my_day = start_date.getDate();
    my_day = prependZeroIfNecessary(my_day);

    var my_hour = start_date.getHours();
    my_hour = prependZeroIfNecessary(my_hour);

    var my_min = start_date.getMinutes();
    my_min = prependZeroIfNecessary(my_min);

    var toReturn = my_year + my_month + my_day + " - " + my_hour + my_min + " - "

    //console.log(toReturn)
    return toReturn

}

function extractDocTitleFromEvent(some_event) {
    // given a calendar event will create a string title for the doc
    // in the form of
    // yyyymmdd - hhmm - title of event
    var eventDetailsArray = getEventDetails(some_event);

    var my_title = parseDatesToFormat(eventDetailsArray[0]);

    my_title = my_title + eventDetailsArray[1];

    return my_title;
}

function makeNotesDoc(some_event) {
    // for a calendar event creates a Google doc for note taking
    // populates doc with attendees and meeting description

    var title = extractDocTitleFromEvent(some_event);

    var doc = DocumentApp.create(title);

    var body = doc.getBody();




    // Create Attendees section
    var header = body.appendParagraph("Attendees")
    header.setHeading(DocumentApp.ParagraphHeading.HEADING2)

    var attendeesArray = getEventDetails(some_event)[3]
    for (y = 0; y < attendeesArray.length; y++) {
        body.appendParagraph(attendeesArray[y].getName());
      }


    // Create Meeting description section
    var header = body.appendParagraph("Meeting Description")
    header.setHeading(DocumentApp.ParagraphHeading.HEADING2)


    var eventDesc = getEventDetails(some_event)[2];
    eventDesc = eventDesc.replace(/<[^>]*>?/gm, '')
    body.appendParagraph(eventDesc)


    // Background section
    var header = body.appendParagraph("Background")
    header.setHeading(DocumentApp.ParagraphHeading.HEADING2)
    body.appendParagraph("...")

    // Main Notes section
    var header = body.appendParagraph("Notes and Comments")
    header.setHeading(DocumentApp.ParagraphHeading.HEADING2)
    body.appendParagraph("...")


    // format text into Montserrat font
    // change here if different font required
    var text = body.editAsText();
    text.setFontFamily("Montserrat")



}

function listAttendees(some_event) {
    // test function
    var attendeesArray = getEventDetails(some_event)[3]
    for (i = 0; i < attendeesArray.length; i++) {
        //body.appendParagraph(attendeesArray[i]);
        console.log(attendeesArray[i].getName())
      }
}

// main driver

function main() {

    


    var my_events = getEventsToday();

    console.log("Number of Events")
    console.log(my_events.length)

    for (z = 0; z < my_events.length; z++) {
        console.log(extractDocTitleFromEvent(my_events[z]))
        makeNotesDoc(my_events[z]);
    }


    var ui = SpreadsheetApp.getUi();
    ui.alert("Documents Created. Process Complete. Documents can be found in MyDrive.")
}