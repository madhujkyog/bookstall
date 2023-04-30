function createCalendarEvent() {

  //Input
  //Pre-requisite
  //1. Add Calendar services in App Script to use https://developers.google.com/apps-script/advanced/calendar?authuser=0
  //2. Select central timezone in AppScript > Settings and Google Sheet > File > Settings
  let bookstallCalendarId = "780d9c73f715c3b3fc6a1967dad50e00893b82648981f899b623e9e211594e4d@group.calendar.google.com";
  let showUIInteraction = true; //Make it false in testing mode
  let headerNames = {
    centerName : "Center",
    centerCoordinatorEmail : "CenterCoordinatorEmail",
    bookstallVolunteerEmails : "BookStallVolunteers",
    meetingStart : "Meeting Start (CDT)",
    meetingEnd : "Meeting End (CDT)",
    meetingDescription : "Meeting Description",
    meetingTitle : "Meeting Title"
  }

  let ui = null;
  if(showUIInteraction)
    ui = SpreadsheetApp.getUi();

  let specificCenterName = "CT";
  if(showUIInteraction)
    specificCenterName = ui.prompt("Please enter center name").getResponseText();


  // Get data from active sheet
  let sheet = SpreadsheetApp.getActiveSheet();
  let dataRange = sheet.getDataRange();
  let dataRangeValues = dataRange.getValues();

  let centerColumnNumber = dataRangeValues[0].indexOf(headerNames.centerName);
  let centerCordinatorEmailColumnNumber = dataRangeValues[0].indexOf(headerNames.centerCoordinatorEmail);
  let bookstallVolunteerEmailsColumnNumber = dataRangeValues[0].indexOf(headerNames.bookstallVolunteerEmails);
  let meetingStartColumnNumber = dataRangeValues[0].indexOf(headerNames.meetingStart);
  let meetingEndColumnNumber = dataRangeValues[0].indexOf(headerNames.meetingEnd);
  let meetingDescriptionColumnNumber = dataRangeValues[0].indexOf(headerNames.meetingDescription);
  let meetingTitleColumnNumber = dataRangeValues[0].indexOf(headerNames.meetingTitle);


  //Validate all header columns
  if(centerColumnNumber == -1 || centerCordinatorEmailColumnNumber == -1
      || bookstallVolunteerEmailsColumnNumber == -1
      || meetingStartColumnNumber == -1 || meetingEndColumnNumber == -1) {
  console.log("One of the required header not found.");
  if(showUIInteraction)
    ui.alert("ERROR : One of the required header not found.");
  return;
  }

  // Validate if center is valid
  let specificCenterRowNumber = -1;
  for(let row = 0 ; row < dataRange.getLastRow() ; row ++) {
      if (dataRangeValues[row][centerColumnNumber] == specificCenterName) {
        specificCenterRowNumber = row;
        break;
      }
  }

if(specificCenterRowNumber == -1) {
  // Center not found
  console.log(specificCenterName + " not found...");
  if(showUIInteraction)
    ui.alert("ERROR : " + specificCenterName + " not found");
  return;
}

// Validate data required to create event
  
let meetingStart =  dataRangeValues[specificCenterRowNumber][meetingStartColumnNumber]; 
let meetingEnd =  dataRangeValues[specificCenterRowNumber][meetingEndColumnNumber];
let centerCoordinatorEmail = dataRangeValues[specificCenterRowNumber][centerCordinatorEmailColumnNumber];
let bookstallVolunteerEmails = dataRangeValues[specificCenterRowNumber][bookstallVolunteerEmailsColumnNumber];
let meetingDescription = dataRangeValues[specificCenterRowNumber][meetingDescriptionColumnNumber];
let meetingTitle = dataRangeValues[specificCenterRowNumber][meetingTitleColumnNumber];


if(!meetingStart || !meetingEnd) {
  console.log("Please Validate Meeting Start and End Date");
  if(showUIInteraction)
    ui.alert("ERROR : Please Validate Meeting Start and End Date");
  return;
}

if(!Date.parse(meetingStart) || !Date.parse(meetingEnd) ) {
  console.log("Please provide valid date as Meeting Start and End Date");
  if(showUIInteraction)
    ui.alert("ERROR : Please provide valid date as Meeting Start and End Date");
  return;
}


if(!meetingTitle ) {
  console.log("Please provide meeting Title");
  if(showUIInteraction)
    ui.alert("ERROR : Please provide meeting title");
  return;
}


if(!centerCoordinatorEmail ) {
  console.log("Please provide center coorinator email");
  if(showUIInteraction)
    ui.alert("ERROR : Please provide center coorinator email");
  return;
}


if(!bookstallVolunteerEmails ) {
  console.log("Please provide book stall volunteer emails");
  if(showUIInteraction)
    ui.alert("ERROR : Please provide book stall volunteer emails");
  return;
}


let calendar = CalendarApp.getCalendarById(bookstallCalendarId);

// check if there is prior event scheduled
let existingNoteStr = dataRange.getCell(specificCenterRowNumber + 1 , meetingDescriptionColumnNumber + 1).getNote();
let existingNote = {};
if(existingNoteStr) {
  try {
    existingNote = JSON.parse(existingNoteStr);
  } catch(e) {
    console.warn(e);
  }
}


let create = true;
let existingEventId = null;
if(existingNote.eventId) {
  console.log("Detected existing event ID " + existingNote.eventId + " : " + existingNote.lastUpdatedDate);
  let existingEvent = Calendar.Events.get(bookstallCalendarId, existingNote.eventId)

  if(existingEvent) {
    console.log("Fetched existing event with id = " + existingEvent.getId() 
      + " : Status = " + existingEvent.status 
      + " : Meeting Start Time = "+ new Date(existingEvent.start.dateTime));

      console.log("Is meeting not in past = "+ (new Date(existingEvent.start.dateTime) > new    Date   ()));

      console.log("Is meeting not cancelled =" + (existingEvent.status != "cancelled"));
  }
 
  if(
    existingEvent && 
    new Date(existingEvent.start.dateTime) > new Date() && // if event is not in past
    existingEvent.status != "cancelled") // if event is not cancelled
  {
      console.log("Found existing meeting. Will update the existing event.");
      create = false; // update existing meeting
      existingEventId = existingEvent.getId()
  }
}

try {
    let eventData = {
      summary : meetingTitle,
      description : meetingDescription,
      start: {
        dateTime: new Date(meetingStart).toISOString()
      },
      end: {
        dateTime: new Date(meetingEnd).toISOString()
      },
      attendees: [
        { email: centerCoordinatorEmail }
      ],
      // Orange background. Use Calendar.Colors.get() for the full list.
      colorId: 6
    };

    let bookstallVolunteerEmailsArr = bookstallVolunteerEmails.split(',');
    bookstallVolunteerEmailsArr.forEach(email => {if(email) eventData.attendees.push({"email" : email})});
    let event = null;
    if (create) {
      event = Calendar.Events.insert(
                    eventData,
                    bookstallCalendarId,
                    { sendUpdates : 'all' },
                    );
      console.log('Successfully inserted event: ' + event.id);
    } else {
      console.log('Updating event : ' + existingEventId + " with " + JSON.stringify(eventData));
      event = Calendar.Events.update(
              eventData,
              bookstallCalendarId,
              existingEventId,
              { sendUpdates : 'all' },
              );
      console.log('Successfully updated event: ' + event.id);

    }

    console.log(event);
    const note = {
      lastUpdatedDate : event.updated ? event.updated : event.created,
      eventId : event.id
    }

    // set event id as comment
    dataRange.getCell(specificCenterRowNumber + 1 , meetingDescriptionColumnNumber + 1)
      .setNote(JSON.stringify(note));


    console.log("Sucessfully scheduled meeting " + new Date(meetingStart));
    if(showUIInteraction)
      ui.alert("Sucessfully scheduled meeting!");
    return;
    
  } catch (e) {
    console.log('Upsert threw an exception: ' + e);
    return;
  }



}
