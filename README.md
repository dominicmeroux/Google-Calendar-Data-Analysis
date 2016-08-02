# Motivation and Background
A university Parking & Transportation department is piloting use of two shared electric bicycles at an off-campus office. Google Calendar is used as a reservation system - the electric bicycles are booked in the same way as a staff member would book a conference room. The following code analyzes these reservations to provide insight into utilization of the e-bikes. 

This code can be generalized to any other asset reserved through a similar system, or modified for analyzing other types of calendar events. 

# Modify both Python (.py) files by adding MySQL database connection information and run from command line
The .ipynb file was the initial development file and is a useful reference. The final files to use are:

1) Weekly_bCal_Reservations.py automatically reads in a specified calendar, pushes any new data out to the MySQL database, and spits out emails within (under current configuration) the last 7 days.

2) bCal_eBikes_Utilization_Report.py does the same thing as the weekly reservations file, and generates an Excel file with a comprehensive asset utilization report. 

"Reading in and Analyzing Calendar Data by Interfacing Between MySQL and Python.ipynb"

This approach is to read in .ics calendar files from Python, and to interface with a MySQL database for data storage
and analysis using SQL queries, which ideal for quickly pulling desired information. The database approach 
is optional as calendar data read in can be brought into a Pandas dataframe or analyzed in another fashion.  

# Solved Challenges
There is a key issue (resolved in this code) I found in working with the iCalendar package (perhaps someone knows of a solution) 
where I would read in data and the UTC timestamp would show up, but with no offset. After much searching online, the solution I 
came accross was to use the _fix_timezone function from the package created by:
https://github.com/oblique63/Python-GoogleCalendarParser/blob/master/calendar_parser.py. I have included their python file,
calendar_parser.py, for use with my solution (They of course deserve full credit for their Python file). 
