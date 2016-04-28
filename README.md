# License
MIT License

Copyright (c) 2016 Dominique Meroux

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

# Reading-In-and-Analyzing-Calendar-Data-by-Interfacing-Between-MySQL-and-Python

"Reading in and Analyzing Calendar Data by Interfacing Between MySQL and Python.ipynb"

This approach is to read in .ics calendar files from Python, and to interface with a MySQL database for data storage
and analysis using SQL queries, which ideal for quickly pulling desired information. The database approach 
is optional as calendar data read in can be brought into a Pandas dataframe or analyzed in another fashion. I opted for
interfacing with a MySQL database because I intend to create a relational database from other aspects of my project. 

# Challenges
There is a key (resovled in this code) issue I found in working with the iCalendar package (perhaps someone knows of a solution) 
where I would read in data and the UTC timestamp would show up, but with no offset. After much searching online, the solution I 
came accross was to use the _fix_timezone function from the package created by:
https://github.com/oblique63/Python-GoogleCalendarParser/blob/master/calendar_parser.py. I have included their python file,
calendar_parser.py, for use with my solution (They of course deserve full credit for their Python file). 

# Future Improvements
This is a first draft, and I have noted planned improvements I intend to push out as I further develop this. These are commented
out at the end of the document. 
