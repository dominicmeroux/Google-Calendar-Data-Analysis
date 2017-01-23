# Motivation and Background
A university Parking & Transportation department uses six shared electric bicycles (e-bikes) at an off-campus office. Google Calendar is used as a reservation system - the electric bicycles are booked in the same way as a staff member would book a conference room. The following code analyzes these reservations to provide insight into utilization of the e-bikes. 

This code can be generalized to any other asset reserved through a similar system, or modified for analyzing other types of calendar events. 

# Using MySQL Database
The .ipynb file was the initial development file and is a useful reference. The final files to use are:

1) 'Past-Week-Reservations.py' automatically reads in a specified calendar (user must add database information and calendar URL as commented in the Python code), pushes any new data out to the MySQL database, and spits out emails within (under current configuration) the last 7 days.

2) 'Utilization-Report-MySQL.py' does the same thing as the weekly reservations file (thus requires the same database and calendar information added to code), and generates an Excel file with a comprehensive asset utilization report. 

NOTE: the MySQL files were originally developed to analyze two e-bikes in the pilot from 2015 - 2016. After the pilot ended, four additional e-bikes were purchased in late 2016, and I developed a variant of these files meant to use SQLite, as this is easier for managers with Windows PC computers to set up and run. 

# Using SQLite Database

The final files are:

1) 'Utilization-Report-SQLite.py' is the file you can download, add in calendar URLs, change line 179 (https://github.com/dominicmeroux/Google-Calendar-Data-Analysis/blob/master/Utilization-Report-SQLite.py#L179), and add / subtract the number of calendars and names. 

2) The Jupyter Notebook variant of this file may be helpful for step-by-step debugging, especially if you want to make a lot of changes (e.g. adding in 10 more calendars, etc.). 

# Solved Challenges
There is a key issue (resolved in this code) I found in working with the iCalendar package (perhaps someone knows of a solution) 
where I would read in data and the UTC timestamp would show up, but with no offset. After much searching online, the solution I 
came accross was to use the _fix_timezone function from the package created by:
https://github.com/oblique63/Python-GoogleCalendarParser/blob/master/calendar_parser.py. I have included their python file,
calendar_parser.py, for use with my solution (They of course deserve full credit for their Python file). 
