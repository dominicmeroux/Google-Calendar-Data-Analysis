from __future__ import print_function
import icalendar
from icalendar import *
from datetime import date, datetime, timedelta
import pickle
import csv
import pandas
from pandas.io import sql
import matplotlib.pyplot as plt
import xlsxwriter
import numpy as np
import sqlite3
import os
import re
import glob
import pytz
from StringIO import StringIO
import calendar_parser as cp 
from urllib import urlopen 
# for calendar_parser, I downloaded the Python file created for this package
# https://github.com/oblique63/Python-GoogleCalendarParser/blob/master/calendar_parser.py
# and saved it in the working directory with my Python file (Jupyter Notebook file). 
# In calendar_parser.py, their function _fix_timezone is very crucial for my code to 
# display the correct local time. 

########################################################
# Establish Connection with SQLite database
########################################################
conn = sqlite3.connect("CalEbikes.db") # or use :memory: to put it in RAM
c = conn.cursor()

###### References and testing
## If table needs to be removed due to error
#c.execute("""DROP TABLE eBike""")
#conn.commit()

## Create table
try:
    c.execute("""CREATE TABLE eBike (eBikeName text, Organizer text, Created text, Start text, End text)""")
    conn.commit()
except:
    print("Table eBike exists")
## Testing input values
#c.execute("INSERT INTO eBike VALUES ('Honeybee', 'dmeroux@yahoo.com', '7/24/2012', '7/24/2012', '7/24/2012')")
#conn.commit()

## Checking what's in the table
#sql = "SELECT * FROM eBike WHERE eBikeName=?"
#c.execute(sql, [("Honeybee")])
#c.fetchall()

# Obtain current count from each calendar to read in and add additional entries only
c.execute("""SELECT COUNT(*) FROM eBike WHERE eBikeName = 'Emerald'""")
EmeraldExistingCount = c.fetchall()

c.execute("""SELECT COUNT(*) FROM eBike WHERE eBikeName = 'Honeybee'""")
HoneybeeExistingCount = c.fetchall()

c.execute("""SELECT COUNT(*) FROM eBike WHERE eBikeName = 'Midnight'""")
MidnightExistingCount = c.fetchall()

c.execute("""SELECT COUNT(*) FROM eBike WHERE eBikeName = 'Smoky'""")
SmokyExistingCount = c.fetchall()

c.execute("""SELECT COUNT(*) FROM eBike WHERE eBikeName = 'Tangerine'""")
TangerineExistingCount = c.fetchall()

c.execute("""SELECT COUNT(*) FROM eBike WHERE eBikeName = 'Violet'""")
VioletExistingCount = c.fetchall()


########################################################
# Read in calendar data
########################################################

# Assign a URL for each calendar
EmeraldURL =    # Calendar 1 URL
HoneybeeURL =   # Calendar 2 URL
MidnightURL =   # Calendar 3 URL
SmokyURL =      # Calendar 4 URL
TangerineURL =  # Calendar 5 URL
VioletURL =     # Calendar 6 URL

# Create list of calendar URLs
URL_list = [EmeraldURL, HoneybeeURL, MidnightURL, SmokyURL, TangerineURL, VioletURL]

# Declare lists
eBikeName = []
Organizer = []
DTcreated = []
DTstart = []
DTend = []
Counter = 0

# For each calendar URL
for i in URL_list:
    Counter = 0
    b = urlopen(i)
    cal = Calendar.from_ical(b.read())
    #timezones = cal.walk('VTIMEZONE')
    #print (timezones)
    #timezones2 = cal.walk('X-WR-TIMEZONE')
    #print(timezones2)
    
    # Obtain length of the calendar
    if (i == EmeraldURL):
        EmeraldLen = len(cal.walk())
    elif (i == HoneybeeURL):
        HoneybeeLen = len(cal.walk())
    elif (i == MidnightURL):
        MidnightLen = len(cal.walk())
    elif (i == SmokyURL):
        SmokyLen = len(cal.walk())
    elif (i == TangerineURL):
        TangerineLen = len(cal.walk())
    elif (i == VioletURL):
        VioletLen = len(cal.walk())
    
    # Read in only new data not previously recorded; read organizer email, reservation creation, start, and end times
    for k in cal.walk():
        if k.name == "VCALENDAR":
            timezone = k.get("X-WR-TIMEZONE")
        if k.name == "VEVENT":
            Counter += 1
            if (i == EmeraldURL):
                if EmeraldLen - Counter > EmeraldExistingCount[0][0]:
                    eBikeName.append('Emerald')
                    Organizer.append( re.sub(r'mailto:', "", str(k.get('ORGANIZER') ) ) )
                    DTcreated.append( cp._fix_timezone( k.decoded('CREATED'), pytz.timezone(timezone) ) )
                    DTstart.append( cp._fix_timezone( k.decoded('DTSTART'), pytz.timezone(timezone) ) )
                    DTend.append( cp._fix_timezone( k.decoded('DTEND'), pytz.timezone(timezone) ) )
            elif (i == HoneybeeURL):
                if HoneybeeLen - Counter > HoneybeeExistingCount[0][0]:
                    eBikeName.append('Honeybee')
                    Organizer.append( re.sub(r'mailto:', "", str(k.get('ORGANIZER') ) ) )
                    DTcreated.append( cp._fix_timezone( k.decoded('CREATED'), pytz.timezone(timezone) ) )
                    DTstart.append( cp._fix_timezone( k.decoded('DTSTART'), pytz.timezone(timezone) ) )
                    DTend.append( cp._fix_timezone( k.decoded('DTEND'), pytz.timezone(timezone) ) )
            elif (i == MidnightURL):
                if MidnightLen - Counter > MidnightExistingCount[0][0]:
                    eBikeName.append('Midnight')
                    Organizer.append( re.sub(r'mailto:', "", str(k.get('ORGANIZER') ) ) )
                    DTcreated.append( cp._fix_timezone( k.decoded('CREATED'), pytz.timezone(timezone) ) )
                    DTstart.append( cp._fix_timezone( k.decoded('DTSTART'), pytz.timezone(timezone) ) )
                    DTend.append( cp._fix_timezone( k.decoded('DTEND'), pytz.timezone(timezone) ) )
            elif (i == SmokyURL):
                if SmokyLen - Counter > SmokyExistingCount[0][0]:
                    eBikeName.append('Smoky')
                    Organizer.append( re.sub(r'mailto:', "", str(k.get('ORGANIZER') ) ) )
                    DTcreated.append( cp._fix_timezone( k.decoded('CREATED'), pytz.timezone(timezone) ) )
                    DTstart.append( cp._fix_timezone( k.decoded('DTSTART'), pytz.timezone(timezone) ) )
                    DTend.append( cp._fix_timezone( k.decoded('DTEND'), pytz.timezone(timezone) ) )
            elif (i == TangerineURL):
                if TangerineLen - Counter > TangerineExistingCount[0][0]:
                    eBikeName.append('Tangerine')
                    Organizer.append( re.sub(r'mailto:', "", str(k.get('ORGANIZER') ) ) )
                    DTcreated.append( cp._fix_timezone( k.decoded('CREATED'), pytz.timezone(timezone) ) )
                    DTstart.append( cp._fix_timezone( k.decoded('DTSTART'), pytz.timezone(timezone) ) )
                    DTend.append( cp._fix_timezone( k.decoded('DTEND'), pytz.timezone(timezone) ) )
            elif (i == VioletURL):
                if VioletLen - Counter > VioletExistingCount[0][0]:
                    eBikeName.append('Violet')
                    Organizer.append( re.sub(r'mailto:', "", str(k.get('ORGANIZER') ) ) )
                    DTcreated.append( cp._fix_timezone( k.decoded('CREATED'), pytz.timezone(timezone) ) )
                    DTstart.append( cp._fix_timezone( k.decoded('DTSTART'), pytz.timezone(timezone) ) )
                    DTend.append( cp._fix_timezone( k.decoded('DTEND'), pytz.timezone(timezone) ) )
b.close()



########################################################
# SQLite and Excel connection for desired data
########################################################

# Now that calendar data is fully read in, create a list with data in a format for 
# entering into the SQLite database
eBikeData = []
for i in range(len(DTcreated)):
    if (Organizer[i] != 'ADMIN_ADDRESS@EMAIL.edu'): # INSERT ANY EMAIL ADDRESS YOU DON'T WANT CONSIDERED FOR RESERVATION DATA
        eBikeData.append((eBikeName[i], Organizer[i], DTcreated[i], DTstart[i], DTend[i]))

# Insert calendar data into SQLite table eBike
c.executemany("INSERT INTO eBike (eBikeName, Organizer, Created, Start, End) VALUES (?, ?, ?, ?, ?)", eBikeData)
conn.commit()

# Find emails associated with reservations created at latest 7 days ago
c.execute("SELECT DISTINCT Organizer FROM eBike WHERE julianday('now') - julianday(Start) <= 7 AND julianday('now') - julianday(Start) >= 0")
WeeklyEmail = c.fetchall()
Email = []
for i in range(len(WeeklyEmail)):
    Email.append(WeeklyEmail[i][0])
    if(Email[i] != 'None'):
        print(Email[i])

# https://xlsxwriter.readthedocs.org
# Workbook Document Name
workbook = xlsxwriter.Workbook('E-BikeUpdate' + datetime.strftime(datetime.now(), "%Y-%m-%d") + '.xlsx')

# Define 'bold' format
bold = workbook.add_format({'bold': True})
format1 = workbook.add_format({'bold': 1,
                               'bg_color': '#3CDAE5',
                               'font_color': '#092A51'})
format2 = workbook.add_format({'bold': 1,
                               'bg_color': '#DA7BD0',
                               'font_color': '#A50202'})

# Add Intro Sheet
worksheet = workbook.add_worksheet('INTRO')
worksheet.write('A1', 'Sheet', bold)
worksheet.write('A2', 'Ebike_Rides_by_User')
worksheet.write('A3', 'Trips_by_Res_Time')
worksheet.write('A4', 'Trips_by_Weekday')
worksheet.write('A5', 'Utilization')
worksheet.write('A6', 'Aggregate_Advance_Reservation')
worksheet.write('A7', 'Time_Series_Advance_Reservation')

worksheet.write('B1', 'Description', bold)
worksheet.write('B2', 'Total E-Bike Rides by User Email')
worksheet.write('B3', 'Total E-Bike Rides by Reservation Hour')
worksheet.write('B4', 'Total E-Bike Rides by Weekday')
worksheet.write('B5', 'Average and Maximum Percent and Hours Utilization')
worksheet.write('B6', 'Number of Days E-Bikes Were Reserved in Advance, by Count of Reservations')
worksheet.write('B7', 'Number of Days E-Bikes Were Reserved in Advance, by Reservation Start Datetime')

c.execute("SELECT Organizer, COUNT(*) AS Total_Rides FROM eBike GROUP BY Organizer ORDER BY Total_Rides DESC;")
TotalRides_by_User = c.fetchall()

# Worksheet Name
worksheet1 = workbook.add_worksheet('Ebike_Rides_by_User')

# Column Names
worksheet1.write('A1', 'User', bold)
worksheet1.write('B1', 'Total Rides', bold)

# Declare Starting Point for row, col
row = 1
col = 0

# Iterate over the data and write it out row by row
for UserEmail, UserRideCount in (TotalRides_by_User):
    worksheet1.write(row, col,     UserEmail)
    worksheet1.write(row, col + 1, UserRideCount)
    row += 1

# Conditional Formatting: E-bike Users with 20+ Rides
worksheet1.conditional_format('B1:B9999', {'type':     'cell',
                                        'criteria': '>=',
                                        'value':    20,
                                        'format':   format1})

c.execute("SELECT strftime('%H', Start) AS Hour_24, strftime('%M', Start) AS Reservation_Time, COUNT(*) AS Total_Rides FROM eBike GROUP BY Reservation_Time, Hour_24 ORDER BY Hour_24 ASC")
Trips_by_Time = c.fetchall()

# Worksheet Name
worksheet2 = workbook.add_worksheet('Trips_by_Res_Time')  # Data.

# Column Names
#worksheet2.write('A1', 'Hour', bold)
worksheet2.write('A1', 'Time', bold)
worksheet2.write('B1', 'Total Rides', bold)

# Declare Starting Point for row, col
row = 1
col = 0

# Iterate over the data and write it out row by row
for Hour_24, Reservation_Time, Total_Rides in (Trips_by_Time):
    worksheet2.write(row, col,     str(Hour_24)+':'+str(Reservation_Time))
    #worksheet2.write(row, col,     Reservation_Time)
    worksheet2.write(row, col + 1, Total_Rides)
    row += 1
    
# Add Chart
chart = workbook.add_chart({'type': 'line'})

# Add Data to Chart
chart.add_series({
    'categories': '=Trips_by_Res_Time!$A$2:$A$16',
    'values':     '=Trips_by_Res_Time!$B$2:$B$16',
    'fill':       {'color': '#791484'},
    'border':     {'color': '#52B7CB'}
})

# Format Chart
chart.set_title({
    'name': 'Total Rides by Reservation Time',
    'name_font': {
        'name': 'Calibri',
        'color': '#52B7CB',
    },
})

chart.set_x_axis({
    'name': 'Reservation Time',
    'empty_cells': 'gaps',
    'name_font': {
        'name': 'Calibri',
        'color': '#52B7CB'
    },
    'num_font': {
        'name': 'Arial',
        'color': '#52B7CB',
    },
})

chart.set_y_axis({
    'name': 'Total Rides',
    'empty_cells': 'gaps',
    'name_font': {
        'name': 'Calibri',
        'color': '#52B7CB'
    },
    'num_font': {
        'italic': True,
        'color': '#52B7CB',
    },
})

# Remove Legend
chart.set_legend({'position': 'none'})

# Insert Chart
worksheet2.insert_chart('E1', chart)

c.execute("SELECT case strftime('%w', Start) when '0' then 'Sunday' when '1' then 'Monday' when '2' then 'Tuesday' when '3' then 'Wednesday' when '4' then 'Thursday' when '5' then 'Friday' when '6' then 'Saturday' else '' end AS Weekday, COUNT(*) AS Total_Rides FROM eBike GROUP BY Weekday ORDER BY Weekday")
Trips_by_Weekday = c.fetchall()

# Worksheet Name
worksheet3 = workbook.add_worksheet('Trips_by_Weekday')

# Column Names
worksheet3.write('A1', 'Weekday', bold)
worksheet3.write('B1', 'Total Rides', bold)

# Declare Starting Point for row, col
row = 1
col = 0

# Iterate over the data and write it out row by row
for Weekday, Total_Rides_by_Weekday in (Trips_by_Weekday):
    worksheet3.write(row, col,     Weekday)
    worksheet3.write(row, col + 1, Total_Rides_by_Weekday)
    row += 1
    
# Add Chart
chart = workbook.add_chart({'type': 'line'})

# Add Data to Chart
chart.add_series({
    'categories': '=Trips_by_Weekday!$A$2:$A$8',
    'values':     '=Trips_by_Weekday!$B$2:$B$8',
    'fill':       {'color': '#791484'},
    'border':     {'color': '#52B7CB'}
})

# Format Chart
chart.set_title({
    'name': 'Total Rides by Weekday',
    'name_font': {
        'name': 'Calibri',
        'color': '#52B7CB',
    },
})

chart.set_x_axis({
    'name': 'Weekday',
    'name_font': {
        'name': 'Calibri',
        'color': '#52B7CB'
    },
    'num_font': {
        'name': 'Arial',
        'color': '#52B7CB',
    },
})

chart.set_y_axis({
    'name': 'Total Rides',
    'name_font': {
        'name': 'Calibri',
        'color': '#52B7CB'
    },
    'num_font': {
        'italic': True,
        'color': '#52B7CB',
    },
})

# Remove Legend
chart.set_legend({'position': 'none'})

# Insert Chart
worksheet3.insert_chart('E1', chart)

c.execute("SELECT case strftime('%w', Start) when '0' then 'Sunday' when '1' then 'Monday' when '2' then 'Tuesday' when '3' then 'Wednesday' when '4' then 'Thursday' when '5' then 'Friday' when '6' then 'Saturday' else '' end AS Weekday, MAX((julianday(End) - julianday(Start))*24) AS Max_Hours, MAX(julianday(End) - julianday(Start))*100 AS Max_PCT_Utilization, AVG((julianday(End) - julianday(Start))*24) AS Avg_Hours, AVG((julianday(End) - julianday(Start))*100) AS Avg_PCT_Utilization FROM eBike GROUP BY Weekday ORDER BY Weekday")
Avg_Max_Hours_PCTutilization_by_Weekday = c.fetchall()

# Worksheet Name
worksheet4 = workbook.add_worksheet('Utilization')

# Column Names
worksheet4.write('A1', 'Weekday', bold)
worksheet4.write('B1', 'Maximum Reservation Duration (hrs)', bold)
worksheet4.write('C1', 'Maximum Percentage Utilization', bold)
worksheet4.write('D1', 'Average Reservation Duration (hrs)', bold)
worksheet4.write('E1', 'Average Percent Utilization', bold)

# Declare Starting Point for row, col
row = 1
col = 0

# Iterate over the data and write it out row by row
for Weekday_AMH, Max_Hours, Max_PCT_Utilization, Avg_Hours, Avg_PCT_Utilization in (Avg_Max_Hours_PCTutilization_by_Weekday):
    worksheet4.write(row, col,     Weekday_AMH)
    worksheet4.write(row, col + 1, Max_Hours)
    worksheet4.write(row, col + 2, Max_PCT_Utilization)
    worksheet4.write(row, col + 3, Avg_Hours)
    worksheet4.write(row, col + 4, Avg_PCT_Utilization)
    row += 1
    
# Conditional Formatting: Percent Utilization Greater Than 50
worksheet4.conditional_format('E2:E8', {'type':     'cell',
                                        'criteria': '>=',
                                        'value':    30,
                                        'format':   format1})

# How far in advance reservations are created
c.execute("SELECT julianday(Start) - julianday(Created) AS Days_Advance_Reservation, COUNT(*) AS Number_Reserved_Trips FROM eBike WHERE julianday(Start) - julianday(Created) >= 0 GROUP BY Days_Advance_Reservation ORDER BY Days_Advance_Reservation DESC")
Advance_Reservation = c.fetchall()

# Worksheet Name
worksheet5 = workbook.add_worksheet('Aggregate_Advance_Reservation')

# Column Names
worksheet5.write('A1', 'Days E-Bike was Reserved Ahead of Time', bold)
worksheet5.write('B1', 'Total Reservations', bold)

# Declare Starting Point for row, col
row = 1
col = 0

# Iterate over the data and write it out row by row
for Days_Advance_Reservation, Number_Reserved_Trips in (Advance_Reservation):
    worksheet5.write(row, col,     Days_Advance_Reservation)
    worksheet5.write(row, col + 1, Number_Reserved_Trips)
    row += 1
    
worksheet5.conditional_format('B2:B9999', {'type':     'cell',
                                        'criteria': '>=',
                                        'value':    5,
                                        'format':   format2})

# Time series of how far in advance reservations are created
c.execute("SELECT Start, julianday(Start) - julianday(Created) AS Days_Advance_Reservation FROM eBike WHERE julianday(Start) - julianday(Created) > 0 ORDER BY Start ASC")
Time_Series_Advance_Reservation = c.fetchall()

Starts = []
for i in range(0, len(Time_Series_Advance_Reservation)): 
    Starts.append(str(Time_Series_Advance_Reservation[i][0]))

# Worksheet Name
worksheet6 = workbook.add_worksheet('Time_Series_Advance_Reservation')

# Column Names
worksheet6.write('A1', 'Reservation Start Date', bold)
worksheet6.write('B1', 'Days E-Bike was Reserved Ahead of Time', bold)

# Declare Starting Point for row, col
row = 1
col = 0

# Iterate over the data and write it out row by row
for StartVal in Starts:
    worksheet6.write(row, col, StartVal)
    row += 1

row = 1
for Start, Days_Advance_Reservation in (Time_Series_Advance_Reservation):
    worksheet6.write(row, col + 1, Days_Advance_Reservation)
    row += 1
    
# Add Chart
chart = workbook.add_chart({'type': 'line'})

worksheet6.conditional_format('B2:B9999', {'type':     'cell',
                                        'criteria': '>=',
                                        'value':    5,
                                        'format':   format2})
workbook.close()
c.close()
conn.close()
