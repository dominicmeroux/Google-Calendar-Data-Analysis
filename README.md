# Reading-In-and-Analyzing-Calendar-Data-by-Interfacing-Between-MySQL-and-Python

"Reading in and Analyzing Calendar Data by Interfacing Between MySQL and Python.ipynb"

This approach is to read in .ics calendar files from Python, and to interface with a MySQL database for data storage
and analysis using SQL queries, which ideal for quickly pulling desired information. The database approach 
is optional as calendar data read in can be brought into a Pandas dataframe or analyzed in another fashion.  

# Challenges
There is a key issue (resolved in this code) I found in working with the iCalendar package (perhaps someone knows of a solution) 
where I would read in data and the UTC timestamp would show up, but with no offset. After much searching online, the solution I 
came accross was to use the _fix_timezone function from the package created by:
https://github.com/oblique63/Python-GoogleCalendarParser/blob/master/calendar_parser.py. I have included their python file,
calendar_parser.py, for use with my solution (They of course deserve full credit for their Python file). 
