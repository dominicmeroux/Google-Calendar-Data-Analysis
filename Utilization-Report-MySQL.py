from __future__ import print_function
from icalendar import *
from datetime import date, datetime, timedelta
import mysql.connector
from mysql.connector import errorcode
import pickle
import csv
import pandas
from pandas.io import sql
import matplotlib.pyplot as plt
import xlsxwriter
import numpy as np
import os
import re
import glob
import pytz
from StringIO import StringIO
#from zipfile import ZipFile
from urllib import urlopen
import calendar_parser as cp 
# for calendar_parser, I downloaded the Python file created for this package
# https://github.com/oblique63/Python-GoogleCalendarParser/blob/master/calendar_parser.py
# and saved it in the working directory with my Python file (Jupyter Notebook file). 
# In calendar_parser.py, their function _fix_timezone is very crucial for my code to 
# display the correct local time. 

USER = # enter database username
PASS = # enter database password
HOST = # enter hostname, e.g. '127.0.0.1'
cnx = mysql.connector.connect(user=USER, password=PASS, host=HOST)

cursor = cnx.cursor()

# Approach / Code modified from MySQL Connector web page
DB_NAME = "CalDb"

# 1) Creates database if it doesn't already exist
# 2) Then connects to the database
def create_database(cursor):
    try:
        cursor.execute(
            "CREATE DATABASE {} DEFAULT CHARACTER SET 'utf8'".format(DB_NAME))
    except mysql.connector.Error as err:
        print("Failed creating database: {}".format(err))
        exit(1)

try:
    cnx.database = DB_NAME    
except mysql.connector.Error as err:
    if err.errno == errorcode.ER_BAD_DB_ERROR:
        create_database(cursor)
        cnx.database = DB_NAME
    else:
        print(err)
        exit(1)

# Create table specifications
TABLES = {}
TABLES['eBike'] = (
    "CREATE TABLE IF NOT EXISTS `eBike` ("
    "  `eBikeName` varchar(10),"
    "  `Organizer` varchar(100),"
    "  `Created` datetime NOT NULL,"
    "  `Start` datetime NOT NULL,"
    "  `End` datetime NOT NULL"
    ") ENGINE=InnoDB")

# If table does not already exist, this code will create it based on specifications
for name, ddl in TABLES.iteritems():
    try:
        print("Creating table {}: ".format(name), end='')
        cursor.execute(ddl)
    except mysql.connector.Error as err:
        if err.errno == errorcode.ER_TABLE_EXISTS_ERROR:
            print("already exists.")
        else:
            print(err.msg)
    else:
        print("OK")

# Obtain current count from each calendar to read in and add additional entries only
cursor.execute("SELECT COUNT(*) FROM eBike WHERE eBikeName = 'Gold'")
GoldExistingCount = cursor.fetchall()

cursor.execute("SELECT COUNT(*) FROM eBike WHERE eBikeName = 'Blue'")
BlueExistingCount = cursor.fetchall()

# Declare lists
eBikeName = []
Organizer = []
DTcreated = []
DTstart = []
DTend = []
Counter = 0

Cal1URL = # Google Calendar URL (from Calendar Settings -> Private Address)
Cal2URL = # URL of second Google Calendar...can scale this code to as many calendars as desired 
          # at an extremily large number (e.g. entire company level), could modify and use parallel processing (e.g. PySpark)

Blue = Cal1URL
Gold = Cal2URL
URL_list = [Blue, Gold]

for i in URL_list:
    Counter = 0
    b = urlopen(i)
    cal = Calendar.from_ical(b.read())
    timezones = cal.walk('VTIMEZONE')
    
    if (i == Blue):
        BlueLen = len(cal.walk())
    elif (i == Gold):
        GoldLen = len(cal.walk())

    #print (cal)
    #print ("Stuff")
    #print (cal.subcomponents)
    for k in cal.walk():
        if k.name == "VEVENT":
            Counter += 1
            if (i == Blue):
                if BlueLen - Counter > GoldExistingCount[0][0]:
                    eBikeName.append('Blue')
                    Organizer.append( re.sub(r'mailto:', "", str(k.get('ORGANIZER') ) ) )
                    DTcreated.append( cp._fix_timezone( k.decoded('CREATED'), pytz.timezone(timezones[0]['TZID']) ) )
                    DTstart.append( cp._fix_timezone( k.decoded('DTSTART'), pytz.timezone(timezones[0]['TZID']) ) )
                    DTend.append( cp._fix_timezone( k.decoded('DTEND'), pytz.timezone(timezones[0]['TZID']) ) )
                    #print (k.property_items('ATTENDEE'))
            elif (i == Gold):
                if GoldLen - Counter > BlueExistingCount[0][0]:
                    eBikeName.append('Gold')
                    Organizer.append( re.sub(r'mailto:', "", str(k.get('ORGANIZER') ) ) )
                    DTcreated.append( cp._fix_timezone( k.decoded('CREATED'), pytz.timezone(timezones[0]['TZID']) ) )
                    DTstart.append( cp._fix_timezone( k.decoded('DTSTART'), pytz.timezone(timezones[0]['TZID']) ) )
                    DTend.append( cp._fix_timezone( k.decoded('DTEND'), pytz.timezone(timezones[0]['TZID']) ) )         
    b.close()

# Now that calendar data is fully read in, create a list with data in a format for 
# entering into the MySQL database. 
# 
# At this point, if the MySQL Connector component is not desired, other approaches  
# include creating a Pandas dataframe or something else.
# For reference, a Pandas dataframe could be created with the following command: 
# df = pandas.DataFrame({'ORGANIZER' : Organizer,'CREATED' : DTcreated, 'DTSTART' : DTstart,'DTEND': DTend})
eBikeData = [] #####################################################
for i in range(len(DTcreated)):
    # Add in condition that the organizer email address cannot be 'none' or any other P&T Management email
    if (Organizer[i] != 'None' and Organizer[i] != 'lauren.bennett@berkeley.edu' and Organizer[i] != 'dmeroux@berkeley.edu' and Organizer[i] != 'berkeley.edu_534da9tjgdsahifulshf42lfbo@group.calendar.google.com'):
        eBikeData.append((eBikeName[i], Organizer[i], DTcreated[i], DTstart[i], DTend[i]))

# Insert calendar data into MySQL table eBike
cursor.executemany("INSERT INTO eBike (eBikeName, Organizer, Created, Start, End) VALUES (%s, %s, %s, %s, %s)", 
                   eBikeData)
cnx.commit()

# Find emails associated with reservations created at latest 6 days ago
cursor.execute("SELECT DISTINCT Organizer FROM eBike WHERE DATEDIFF(CURDATE(), Start) <= 6 AND DATEDIFF(CURDATE(), Start) >= 0")
WeeklyEmail = cursor.fetchall()
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

### Total e-Bike Rides by User
cursor.execute("SELECT Organizer, COUNT(*) AS Total_Rides FROM eBike GROUP BY Organizer ORDER BY Total_Rides DESC;")
TotalRides_by_User = cursor.fetchall()

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

### Total Trips by Reservation Time
cursor.execute("SELECT EXTRACT(HOUR FROM Start) AS Hour_24, DATE_FORMAT(Start, '%h %p') AS Reservation_Time, COUNT(*) AS Total_Rides FROM eBike GROUP BY Reservation_Time, Hour_24 ORDER BY Hour_24 ASC")
Trips_by_Time = cursor.fetchall()

# Worksheet Name
worksheet2 = workbook.add_worksheet('Trips_by_Res_Time')  # Data.

# Column Names
worksheet2.write('A1', 'Reservation Start Time', bold)
worksheet2.write('B1', 'Total Rides', bold)

# Declare Starting Point for row, col
row = 1
col = 0

# Iterate over the data and write it out row by row
for Hour_24, Reservation_Time, Total_Rides in (Trips_by_Time):
    worksheet2.write(row, col,     Reservation_Time)
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
    'name': 'Total Rides by Reservation Start Time',
    'name_font': {
        'name': 'Calibri',
        'color': '#52B7CB',
    },
})

chart.set_x_axis({
    'name': 'Reservation Start Time',
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

# GO TO END OF DATA

### Total Trips by Weekday
cursor.execute("SELECT DAYNAME(Start) AS Weekday, COUNT(*) AS Total_Rides FROM eBike GROUP BY Weekday ORDER BY FIELD(Weekday, 'MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY', 'SUNDAY')")
Trips_by_Weekday = cursor.fetchall()

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
    'categories': '=Trips_by_Weekday!$A$2:$A$8)',
    'values':     '=Trips_by_Weekday!$B$2:$B$8)',
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

### Average and Maximum Hours and Percent Utilization by Weekday

cursor.execute("SELECT DAYNAME(Start) AS Weekday, MAX((HOUR(End - Start)*60 + MINUTE(End - Start))/60) AS Max_Hours, (MAX((HOUR(End - Start)*60 + MINUTE(End - Start))/60)/8)*100 AS Max_PCT_Utilization, AVG((HOUR(End - Start)*60 + MINUTE(End - Start))/60) AS Avg_Hours, (AVG((HOUR(End - Start)*60 + MINUTE(End - Start))/60)/8)*100 AS Avg_PCT_Utilization FROM eBike WHERE (((HOUR(End - Start)*60 + MINUTE(End - Start))/60)/8)*100 < 95 GROUP BY Weekday ORDER BY FIELD(Weekday, 'MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY', 'SUNDAY')")
Avg_Max_Hours_PCTutilization_by_Weekday = cursor.fetchall()

# Worksheet Name
worksheet4 = workbook.add_worksheet('Utilization')

# Column Names
worksheet4.write('A1', 'Weekday', bold)
worksheet4.write('B1', 'Maximum Reservation Duration (hrs)', bold)
worksheet4.write('C1', 'Maximum Percentage Utilization', bold)
worksheet4.write('D1', 'Average Reservation Duration (hrs)', bold)
worksheet4.write('E1', 'Average Percent Utilization', bold)
worksheet4.write('F1', 'NOTE: A small handfull of outliers above 95% utilization are excluded', bold)

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

############################################
cursor.execute("SELECT Start, End, DAYNAME(Start) AS Weekday, ((HOUR(End - Start)*60 + MINUTE(End - Start))/60) AS Hours, (((HOUR(End - Start)*60 + MINUTE(End - Start))/60)/8)*100 AS PCT_Utilization FROM eBike ORDER BY (((HOUR(End - Start)*60 + MINUTE(End - Start))/60)/8)*100 DESC")
Utilization = cursor.fetchall()

worksheet4.write('A11', 'Reservation Start', bold)
worksheet4.write('B11', 'Reservation End', bold)
worksheet4.write('C11', 'Weekday', bold)
worksheet4.write('D11', 'Hours Reserved', bold)
worksheet4.write('E11', 'Percent Utilization', bold)

row += 3
col = 0
count = 12
for Start, End, Day, Hour, PCT_Utilization in (Utilization):
    worksheet4.write(row, col, Start) ########################## https://xlsxwriter.readthedocs.io/working_with_dates_and_time.html
    worksheet4.write(row, col + 1, End) #####
    worksheet4.write(row, col + 2, Day) #####
    worksheet4.write(row, col + 3, Hour)
    worksheet4.write(row, col + 4, PCT_Utilization)
    row += 1
    if (PCT_Utilization > 95.0):
        count += 1

# Add Chart
chart = workbook.add_chart({'type': 'column'})

# Add Data to Chart
chart.add_series({
    'values': '=Utilization!$E$'+str(count)+':$E$'+str(len(Utilization)),
    'fill':   {'color': '#52B7CB'},
    'border': {'color': '#52B7CB'}
})

count = 0

# Format Chart
chart.set_title({
    'name': 'Percent Utilization',
    'name_font': {
        'name': 'Calibri',
        'color': '#52B7CB',
    },
})

chart.set_x_axis({
    'name': 'Reservation',
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
    'name': 'Percent Utilization',
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
worksheet4.insert_chart('G4', chart)
####


### How far in advance reservations are created
# How far in advance reservations are created
cursor.execute("SELECT DATEDIFF(Start, Created) AS Days_Advance_Reservation, COUNT(*) AS Number_Reserved_Trips FROM eBike WHERE DATEDIFF(Start, Created) >= 0 GROUP BY Days_Advance_Reservation ORDER BY Days_Advance_Reservation DESC")
Advance_Reservation = cursor.fetchall()

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
cursor.execute("SELECT Start, DATEDIFF(Start, Created) AS Days_Advance_Reservation FROM eBike WHERE DATEDIFF(Start, Created) > 0 ORDER BY Start ASC")
Time_Series_Advance_Reservation = cursor.fetchall()

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
cursor.close()
cnx.close()
