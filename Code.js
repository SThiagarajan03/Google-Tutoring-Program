// The function below creates a Google Forms for students to fill out which will be used to schedule an appointment
// Sources Used: 1,2,3,4,10,11,25,30
function createGoogleForms() { 
  // Create a Google Forms
  let googleForm = FormApp.create('Bangor Tutoring');

  // Add several questions to collect information on the students
  googleForm.addTextItem()
  .setTitle('What is your name?');

  googleForm.addMultipleChoiceItem()
  .setTitle('What school do you go to?')
  .setChoiceValues(['School 1','School 2', 'School 3'])

  googleForm.addTextItem()
  .setTitle('What grade are you in?');

  googleForm.addDateTimeItem()
  .setTitle('When would you like to be tutored?')

  googleForm.addTextItem()
  .setTitle('What is your email address?');

  let studentSubject = googleForm.addMultipleChoiceItem()
  .setTitle('What subject would you like to be tutored in?')

  // Add sections, which create subsets of the form which the user goes to based on their answers to certain questions.
  // In this case, if the user selects that they would like to be tutored in "Math", they will go to the sub-form including the math tutors.
  let mathSection = googleForm.addPageBreakItem()
  .setTitle('Math')
  .setGoToPage(FormApp.PageNavigationType.SUBMIT); // This line allows the form to be submitted after completing this section

  googleForm.addMultipleChoiceItem()
  .setTitle('Who would you like to be tutored by?')
  .setChoiceValues(['Tutor 1', 'Tutor 2']);

  googleForm.addParagraphTextItem()
  .setTitle('Please provide some information on what you would like to be tutored in.');

  // Create a section and the same questions for the math section for each of the other subjects
  let scienceSection = googleForm.addPageBreakItem()
  .setTitle('Science')
  .setGoToPage(FormApp.PageNavigationType.SUBMIT);

  googleForm.addMultipleChoiceItem()
  .setTitle('Who would you like to be tutored by?')
  .setChoiceValues(['Tutor 1', 'Tutor 2']);

  googleForm.addParagraphTextItem()
  .setTitle('Please provide some information on what you would like to be tutored in.');

  let compSciSection = googleForm.addPageBreakItem()
  .setTitle('Computer Science')
  .setGoToPage(FormApp.PageNavigationType.SUBMIT);

  googleForm.addMultipleChoiceItem()
  .setTitle('Who would you like to be tutored by?')
  .setChoiceValues(['Tutor 1', 'Tutor 2']);

  googleForm.addParagraphTextItem()
  .setTitle('Please provide some information on what you would like to be tutored in.');

  let englishSection = googleForm.addPageBreakItem()
  .setTitle('English')
  .setGoToPage(FormApp.PageNavigationType.SUBMIT);

  googleForm.addMultipleChoiceItem()
  .setTitle('Who would you like to be tutored by?')
  .setChoiceValues(['Tutor 1', 'Tutor 2']);

  googleForm.addParagraphTextItem()
  .setTitle('Please provide some information on what you would like to be tutored in.');

  let historySSSection = googleForm.addPageBreakItem()
  .setTitle('History/Social Studies')
  .setGoToPage(FormApp.PageNavigationType.SUBMIT);

  googleForm.addMultipleChoiceItem()
  .setTitle('Who would you like to be tutored by?')
  .setChoiceValues(['Tutor 1', 'Tutor 2']);

  googleForm.addParagraphTextItem()
  .setTitle('Please provide some information on what you would like to be tutored in.');

  let languagesSection = googleForm.addPageBreakItem()
  .setTitle('Languages')
  .setGoToPage(FormApp.PageNavigationType.SUBMIT);

  googleForm.addMultipleChoiceItem()
  .setTitle('Who would you like to be tutored by?')
  .setChoiceValues(['Tutor 1', 'Tutor 2']);

  googleForm.addParagraphTextItem()
  .setTitle('Please provide some information on what you would like to be tutored in.');

  // Send the user to the appropriate section based on their selected tutoring subject
  studentSubject.setChoices([
    studentSubject.createChoice('Math', mathSection),
    studentSubject.createChoice('Science', scienceSection),
    studentSubject.createChoice('Computer Science', compSciSection),
    studentSubject.createChoice('English', englishSection),
    studentSubject.createChoice('History/Social Studies', historySSSection),
    studentSubject.createChoice('Languages', languagesSection)
  ]);

  // Display the URLs to the form
  Logger.log('URL for filling out the form: ' + googleForm.getPublishedUrl());
  Logger.log('URL for editing: ' + googleForm.getEditUrl());
}

// This function collects the results from the Google Forms and puts them into an easy-to-access list format
// Sources Used: 5,23,24,28,29
function getResults(formId, spreadsheetId) {
  // Access the Google Form
  let googleForm = FormApp.openById(formId);

  // Make a list for all information from each user that filled out the Google Form
  let userAllInfo = [];

  // Loop through the responses and collect the values
  let googleFormResponses = googleForm.getResponses();
  for (let i = 0; i < googleFormResponses.length; i++) {
    let questionResponses = googleFormResponses[i].getItemResponses();

    // Create a list for individual users that will be looped through and overwritten each time
    let userInfo = [];

    for (let j = 0; j < questionResponses.length; j++) {
      // Add the values to the list
      userInfo.push(questionResponses[j].getResponse());
    }

    // Use the tutor's name to find their email based on an already existing spreadsheet which has this information
    // Uses the function getTutorEmail (defined below)
    let tutorName = userInfo[6];
    let tutorEmail = getTutorEmail(tutorName, spreadsheetId)[0]; // Need to do [0] because it is another list
    userInfo.push(tutorEmail)

    // Add the list to the "master" list (userAllInfo)
    userAllInfo.push(userInfo);
  }

  googleForm.deleteAllResponses() // Delete all responses so when program is run again there is no repeat of info

  return userAllInfo // Format: Student Name, School, Grade, Time, Email, Subject, Tutor Name, Appointment Info., Tutor Email
}

// The purpose of this function is to identify a tutor's email based on their name and a spreadsheet which has this information
// Sources Used: 5,6,7,27
function getTutorEmail(tutorName, spreadsheetId) {
  // Open spreadsheet and access the appropriate sheet within (there's only one)
  let tutorsSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
  let tutorsSheet = tutorsSpreadsheet.getSheets()[0];

  // Get the values for column A which contains the names of the tutors
  let tutorNames = tutorsSheet.getRange('A:A').getValues();
  let tutorEmail = '';

  // Loop through column A to find the row corresponding to the tutor's name, then go over to column B to find the email
  for (let i = 0; i < tutorNames.length; i++) {
    if (tutorNames[i] == tutorName) {
      tutorEmail = tutorsSheet.getRange('B' + (i+1).toString()).getValues()[0]; // Have to do [0] because it is returned as a list
    }
  }

  return tutorEmail
}

// This function sends out an email to the student and tutor providing details about the appointment (and acting as a reminder)
// Source Used: 9
function emailSend(results) { // results is returned from getResults()
  for (let i = 0; i < results.length; i++) {
    let userInfo = results[i];

    // Send email to tutor about appointment
    let tutorEmailMsg1 = 'Hello, ' + userInfo[6] + '! ';
    let tutorEmailMsg2 = userInfo[0] + ' wants to be tutored by you in ' + userInfo[5] + '. ';
    let tutorEmailMsg3 = userInfo[0] + ' is in grade ' + userInfo[2] + ' and goes to ' + userInfo[1] + '. ';
    let tutorEmailMsg4 = 'Your appointment is scheduled for ' + userInfo[3] + '. ';
    let tutorEmailMsg5 = 'Here is some information they added about their appointment: ' + userInfo[7] + '.';
    let tutorEmailSubject = 'Your Tutoring Appointment!';
    let tutorEmailMessage = tutorEmailMsg1 + tutorEmailMsg2 + tutorEmailMsg3 + tutorEmailMsg4 + tutorEmailMsg5;
    MailApp.sendEmail(userInfo[8], tutorEmailSubject, tutorEmailMessage);

    // Send email to student about appointment
    let studentEmailMsg1 = 'Hello, ' + userInfo[0] + '! ';
    let studentEmailMsg2 = 'You are going to be tutored by ' + userInfo[6] + '! ';
    let studentEmailMsg3 = 'Your appointment is scheduled for ' + userInfo[3] + '. ';
    let studentEmailMsg4 = 'Here is the information you added about your appointment: ' + userInfo[7] + '.';
    let studentEmailSubject = 'Your Tutoring Appointment!';
    let studentEmailMessage = studentEmailMsg1 + studentEmailMsg2 + studentEmailMsg3 + studentEmailMsg4;
    MailApp.sendEmail(userInfo[4], studentEmailSubject, studentEmailMessage);
  }
}

// This function updates an availabilities spreadsheet to indicate an appointment has been made. It does this by coloring a specific cell red.
// Sources Used: 5,6,7,12,13,14,16,27
function updateSpreadsheet(results, spreadsheetId) { // results is returned from getResults()
  let days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];

  // Open spreadsheet
  let availabilitiesSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
  for (i = 0; i < results.length; i++) {
    // Use the responses from the student to find the day of the week they selected for the appointment
    let appointmentTimeInfo = results[i][3]; // Format: '2021-02-04 15:00'
    let appointmentDate = new Date(appointmentTimeInfo);
    let appointmentDay = appointmentDate.getDay() - 1; // Subtract 1 because it starts week at Sunday
    
    // Go to the corresponding sheet on the availabilities spreadsheet
    let sheetName = days[appointmentDay];
    let availabilitiesSheet = availabilitiesSpreadsheet.getSheetByName(sheetName);

    // Go to the right column based on the time (time is along the top)
    let appointmentTime = parseInt(appointmentTimeInfo.substring(11,13)) - 15; // Subtract 15 to account for time past noon
    let columnns = ['B', 'C', 'D', 'E'];
    let column = columnns[appointmentTime];
    let row = 0;

    // Go to the right row (based on the tutor name, as names are down the left side)
    let tutorName = results[i][6];
    let tutorNames = availabilitiesSheet.getRange('A:A').getValues();
    for (j = 0; j < tutorNames.length; j++) {
      if (tutorNames[j] == tutorName) {
        row = (j + 1).toString(); // Rows start at 1
      }
    }

    // Color in the correct cell red
    let appointmentCell = availabilitiesSheet.getRange(column + row).setBackground('red');
  }
}

// This function creates a 6 digit random ID for the calendar appointment
// Source Used: 19
function generateId() {
  let alphabet = 'abcdefghijklmnopqrstuvwxyz';
  let id1 = alphabet[Math.floor(Math.random() * 26)];
  let id2 = alphabet[Math.floor(Math.random() * 26)];
  let id3 = alphabet[Math.floor(Math.random() * 26)];
  let id4 = Math.floor(Math.random() * 10);
  let id5 = Math.floor(Math.random() * 10);
  let id6 = Math.floor(Math.random() * 10);
  return id1 + id2 + id3 + id4 + id5 + id6
}

// This function makes a Google Calendar appointment in a given Google Calendar
// Sources Used: 5,14,15,18,27
function calendarAppointment(results, calendarId) { // results is returned from getResults()
  for (i = 0; i < results.length; i++) {
    // Set start and end times based on results inputted to Google Forms
    let time = results[i][3]; // Format: '2021-02-04 15:00'
    let timeStart = time.substring(0,10) + 'T' + time.substring(11,16) + ':00'; // Format: '2021-02-04T15:00'
    let timeEnd = time.substring(0,10) + 'T' + time[11] + (parseInt(time[12]) + 1).toString() + ':00:00'; // 1 hour after start
    
    // Create the Google Calendar event
    let event = {
        "summary": 'This is the Google Calendar event for your tutoring appointment.',
       "start": {
         "dateTime": timeStart,
          "timeZone": "America/New_York"
           },
        "end": {
          "dateTime": timeEnd,
          "timeZone": "America/New_York",
            },
             "conferenceData": {
               "createRequest": {
                "conferenceSolutionKey": {
                  "type": "hangoutsMeet"
               },   
                "requestId": generateId()
                }
          }
        };
    
    // Add the event to the Google Calendar
    event = Calendar.Events.insert(event, calendarId, {"conferenceDataVersion": 1});
    
    // Display Google Meet link
    Logger.log(event.hangoutLink)
  }
}

function programTest() {
  let appointmentFormId = 'FORM ID FOR GOOGLE FORM USED TO SCHEDULE APPOINTMENT';
  let tutorEmailSpreadsheetId = 'SPREADSHEET ID FOR SPREADSHEET WITH TUTORS AND CORRESPONDING EMAILS';
  let schedulingSpreadsheetId = 'SPREADSHEET ID FOR SPREADSHEET THAT INDICATES APPOINTMENT TIME SLOTS';
  let calendarId = 'CALENDAR ID FOR SHARED CALENDAR';

  let results = getResults(appointmentFormId, tutorEmailSpreadsheetId);
  calendarAppointment(results, calendarId)
  emailSend(results)
  updateSpreadsheet(results, schedulingSpreadsheetId)
}


