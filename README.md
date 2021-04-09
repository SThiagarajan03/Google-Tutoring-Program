# Google-Tutoring-Program
A tutoring program was developed in G Suite to facilitate connections between students and tutors.

The goal of this project was to create a platform that will allow students easy access to tutors to make for academic gains due to COVID-19 learning losses. Here is the process the tutoring program follows:
1. A student fills out a Google Forms, detailing information about themselves, their school, what they want to be tutored in, and who they want to be tutored by.
2. The program sends out an email to the student and tutor, summarizing the appointment and acting as a reminder.
3. The program creates a Google Calendar appointment at the specified time with an attached Google Meet link in shared Google Calendar.
4. The program updates a Google Sheet integrated within a Google Sites so students do not double-book the same appointment time slot.

Each step after step 1 is automated; that is, the program takes over and does each step on its own. A Google Sites already exists with steps 1 and 4 integrated into it, with the goal of integrating step 3 into it as well. Obviously, step 2 cannot be integrated as it is within an email address.

The code (written in JavaScript, through the Apps Script Editor) has functions for each step, along with a few extra functions for accomplishing sub-tasks. Their purpose will become obvious once the code is viewed.

There is a another document titled "SOURCES" which contains links to all documentation and websites that were used to assist with code development. In addition, each function in the actually code has a preliminary comment explaining what specific sources were used to make that function.
