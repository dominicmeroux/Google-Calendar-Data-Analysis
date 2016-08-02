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
eBikeData = []
for i in range(len(DTcreated)):
    eBikeData.append((eBikeName[i], Organizer[i], DTcreated[i], DTstart[i], DTend[i]))

# Insert calendar data into MySQL table eBike
cursor.executemany("INSERT INTO eBike (eBikeName, Organizer, Created, Start, End) VALUES (%s, %s, %s, %s, %s)", 
                   eBikeData)
cnx.commit()

# Find emails associated with reservations created at latest 7 days ago
cursor.execute("SELECT DISTINCT Organizer FROM eBike WHERE DATEDIFF(CURDATE(), Start) <= 6 AND DATEDIFF(CURDATE(), Start) >= 0")
WeeklyEmail = cursor.fetchall()
Email = []
for i in range(len(WeeklyEmail)):
    Email.append(WeeklyEmail[i][0])
    if(Email[i] != 'None'):
        print(Email[i])
