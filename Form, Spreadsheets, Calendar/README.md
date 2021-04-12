In this folder, there are two images, "Spreadsheet with Tutors and Emails" and "Spreadsheet with Appointment Time Slots". These images explain the structure of the spreadsheets so that the program can work using the data in them. If they don't follow the format in the images, the program will most likely not work. However, the program can quite easily be modified so it accepts data from the spreadsheets in a different format.

In order to use the program, it is necessary to setup spreadsheets using the formats in the images. Once this has been done, it is necessary to get the Spreadsheet IDs. This can be taken from the web address of the spreadsheet (from the EDIT version, not the VIEW version. You can be sure you have the EDIT version by seeing if "edit" is in the web address). See the example below.

Web Address: https://docs.google.com/spreadsheets/d/1234567890abcdefghijklmnopqrstuvwxyz/edit#gid=0

Spreadsheet ID: 1234567890abcdefghijklmnopqrstuvwxyz

It's also important to have Form ID for the program to run as intended. The first function in the code creates a Google Forms with appropriate questions. The Form ID can be found using the example above and the web address of the form (again, make sure to use the EDIT version of the form, not the version that allows users to complete the questions).

The final necessary component is to have a Calendar ID. The best way to get this is to create a new Calendar specifically for the program. This can easily be done by going to Google Calendar. When creating the Calendar, it is important to make sure the Calendar is PUBLIC, otherwise only the creator of the calendar will be able to see the events created in the calendar. To get the Calendar ID, hover over the calendar, click on the 3 dots, and then click on "Settings and sharing". Scroll down the page and you will find the Calendar ID under "Integrate calendar". The ID will end in "group.calendar.google.com".

After getting all the necessary IDs, they can be inputted into the last function in the code, in the appropriate places, which are denoted in the code and can easily be found.
