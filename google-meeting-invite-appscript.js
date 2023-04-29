function createCalendarEvent() {

  // Demo Video - https://drive.google.com/drive/u/0/folders/1VRkh0GJDuoZleFLrDgt9bLX20C8Ria2r
  //Input
  // TODO: Select central timezone in AppScript > Settings and Google Sheet > File > Settings
  let bookstallCalendarId = "780d9c73f715c3b3fc6a1967dad50e00893b82648981f899b623e9e211594e4d@group.calendar.google.com";
  let showUIInteraction = true; //Make it false in testing mode
  let headerNames = {
    centerName : "Center",
    centerCoordinatorEmail : "CenterCoordinatorEmail",
    bookstallVolunteerEmails : "BookStallVolunteers",
    meetingStart : "Meeting Start (CDT)",
    meetingEnd : "Meeting End (CDT)",
    meetingDescription : "Meeting Description",

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
let eventTitle = "BookStall Meeting - " + specificCenterName;
  
let meetingStart =  dataRangeValues[specificCenterRowNumber][meetingStartColumnNumber]; 
let meetingEnd =  dataRangeValues[specificCenterRowNumber][meetingEndColumnNumber];
let centerCoordinatorEmail = dataRangeValues[specificCenterRowNumber][centerCordinatorEmailColumnNumber];
let bookstallVolunteerEmails = dataRangeValues[specificCenterRowNumber][bookstallVolunteerEmailsColumnNumber];
let meetingDescription = dataRangeValues[specificCenterRowNumber][meetingDescriptionColumnNumber];


if(!meetingStart || !meetingEnd) {
  console.log("Please Validate Meeting Start and End Date");
  if(showUIInteraction)
    ui.alert("ERROR : Please Validate Meeting Start and End Date");
  return;
}

if(!meetingStart || !meetingEnd ) {
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

  console.log("Detected existing event ID " + existingNote.eventId);

let foundEvent = null;
if(existingNote.eventId) {
  let event = calendar.getEventById(existingNote.eventId);
  if(event && event.getStartTime() > new Date()) {
      console.log("Found existing meeting. Will update the existing event " + event.getId() + " scheduled on " + event.getStartTime());
      foundEvent = event;
  }
}

console.log("Meeting Start " + meetingStart);

event = foundEvent || calendar.createEvent(eventTitle, new Date(meetingStart) , new Date(meetingEnd));

event.setDescription(meetingDescription)
event.setTime(new Date(meetingStart) , new Date(meetingEnd));
event.addGuest(centerCoordinatorEmail);
event.setGuestsCanInviteOthers(true);

event.addEmailReminder(60); // reminder 60 mins before 

let bookstallVolunteerEmailsArr = bookstallVolunteerEmails.split(',');
bookstallVolunteerEmailsArr.forEach(email => {if(email) event.addGuest(email)});

const note = {
  scheduledDate : new Date().toDateString(),
  eventId : event.getId()
}

// set event id as comment
dataRange.getCell(specificCenterRowNumber + 1 , meetingDescriptionColumnNumber + 1)
  .setNote(JSON.stringify(note));


console.log("Sucessfully scheduled meeting " + new Date(meetingStart));
if(showUIInteraction)
  ui.alert("Sucessfully scheduled meeting!");
return;

}
