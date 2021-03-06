#!/usr/bin/python
# -*- coding:utf-8 -*-
from Tools import tools_v000 as tools
import os
from os.path import dirname
import selenium
from selenium.webdriver.common.keys import Keys
import xlsxwriter
import datetime

# -24 for the name of this project myProtimeExtraction_2020
# save_path = dirname(__file__)[ : -24]
save_path = os.path.dirname(os.path.abspath("__file__"))[ : -24]
print ("save_path = " + save_path)
propertiesFolder_path = save_path + "/"+ "Properties"

# Properties
email_text = tools.readProperty(propertiesFolder_path, 'myProtimeExtraction_2020', 'Email=')
password_text = tools.readProperty(propertiesFolder_path, 'myProtimeExtraction_2020', 'Password=')
excel_path = tools.readProperty(propertiesFolder_path, 'myProtimeExtraction_2020', 'excel_path=')
name_of_file = tools.readProperty(propertiesFolder_path, 'myProtimeExtraction_2020', 'name_of_file=')
year_to_check = tools.readProperty(propertiesFolder_path, 'myProtimeExtraction_2020', 'year_to_check=')
begin_month_date = tools.readProperty(propertiesFolder_path, 'myProtimeExtraction_2020', 'begin_month_date=')
end_month_date = tools.readProperty(propertiesFolder_path, 'myProtimeExtraction_2020', 'end_month_date=')

# static variable
workbook = ''
worksheet = ''
debug = False # Default value is False

# Open Browser
def openBrowser() :
    tools.openBrowserChrome()

def closeBrowser() :
    tools.closeBrowserChrome()

# Start myProtime
def connectToMyProtime() :
    tools.driver.get('https://nn.myprotime.eu/')

def enterCredentials() :
    tools.waitLoadingPageByID2(10, 'Email')
    username = tools.driver.find_element_by_id("Email")
    username.send_keys(email_text)

    tools.waitLoadingPageByID2(10,'Password')
    password = tools.driver.find_element_by_id("Password")
    password.send_keys(password_text)
    password.send_keys(Keys.ENTER)

def goToMyCalendar() :
    tools.waitLoadingPageByXPATH2(10,'//*[@id="scrollzone-wrapper"]/div[2]/button')
    tools.driver.get('https://nn.myprotime.eu/#/calendar/my-calendar')

def goToMyCalendar_specific_date(date) :
    tools.driver.get('https://nn.myprotime.eu/#/calendar/person/113864/month/daydetail/113864/'+date+'?date=2020-01-01')

def recoverInformation() :
    tools.waitLoadingPageByXPATH2(30, '/html/body/div[2]/div/div[1]/div/div[1]/div/div[2]/div')
    tools.waitLoadingPageByXPATH2(30, '/html/body/div[2]/div/div[1]/div/div[1]/div/div[2]/div/div[1]/div[2]/div/div[2]/div[2]')

    try :
        in_1 = tools.driver.find_element_by_xpath("/html/body/div[2]/div/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div[2]/ul/li[1]/a/span[1]/span[2]").text
    except selenium.common.exceptions.NoSuchElementException:
        # When hald day work
        try : 
            in_1 = tools.driver.find_element_by_xpath("/html/body/div[2]/div/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div[3]/ul/li[1]/a/span[1]/span[2]").text
        except selenium.common.exceptions.NoSuchElementException:
            in_1 = ''
    if (debug) :
        print ("in_1       : " + in_1)
    try :
        out_1 = tools.driver.find_element_by_xpath("/html/body/div[2]/div/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div[2]/ul/li[2]/a/span[1]/span[2]").text
    except selenium.common.exceptions.NoSuchElementException:
        try :
            out_1 = tools.driver.find_element_by_xpath("/html/body/div[2]/div/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div[3]/ul/li[2]/a/span[1]/span[2]").text
        except selenium.common.exceptions.NoSuchElementException:
            out_1 = ''
    if (debug) :
        print ("out_1      : " + out_1)
    try :
        in_2 =  tools.driver.find_element_by_xpath("/html/body/div[2]/div/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div[2]/ul/li[3]/a/span[1]/span[2]").text
    except selenium.common.exceptions.NoSuchElementException:
        try :
            in_2 =  tools.driver.find_element_by_xpath("/html/body/div[2]/div/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div[3]/ul/li[3]/a/span[1]/span[2]").text
        except selenium.common.exceptions.NoSuchElementException:
            in_2 = ''
    if (debug) :
        print ("in_2       : " + in_2)
    try :
        out_2 = tools.driver.find_element_by_xpath("/html/body/div[2]/div/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div[2]/ul/li[4]/a/span[1]/span[2]").text
    except selenium.common.exceptions.NoSuchElementException:
        try :
            out_2 = tools.driver.find_element_by_xpath("/html/body/div[2]/div/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div[3]/ul/li[4]/a/span[1]/span[2]").text
        except selenium.common.exceptions.NoSuchElementException:
            out_2 = ''
    if (debug) :
        print ("out_2      : " + out_2)
    try :
        total_hour = tools.driver.find_element_by_xpath("/html/body/div[2]/div/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div[3]/div[2]/ul[1]/li[3]/span[2]").text
        time_prested = tools.driver.find_element_by_xpath("/html/body/div[2]/div/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div[3]/div[2]/ul[1]/li[2]/span[2]").text
    except selenium.common.exceptions.NoSuchElementException:
        # When half day
        try :
            total_hour = tools.driver.find_element_by_xpath("/html/body/div[2]/div/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div[4]/div[2]/ul[1]/li[3]/span[2]").text
            time_prested = tools.driver.find_element_by_xpath("/html/body/div[2]/div/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div[4]/div[2]/ul[1]/li[2]/span[2]").text
        except selenium.common.exceptions.NoSuchElementException:
            try :
                total_hour = tools.driver.find_element_by_xpath("/html/body/div[2]/div/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div[3]/div[2]/ul[1]/li[3]/span[2]").text
                time_prested = tools.driver.find_element_by_xpath("/html/body/div[2]/div/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div[3]/div[2]/ul[1]/li[2]/span[2]").text
            except selenium.common.exceptions.NoSuchElementException:
                #  Exception when it's a day of weekend 
                try :
                    total_hour = tools.driver.find_element_by_xpath("/html/body/div[2]/div/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div[2]/div[2]/ul[1]/li[3]/span[2]").text
                    time_prested = tools.driver.find_element_by_xpath("/html/body/div[2]/div/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div[2]/div[2]/ul[1]/li[2]/span[2]").text
                except selenium.common.exceptions.NoSuchElementException:
                    total_hour = '00:00'
                    time_prested = '00:00'
    if (debug) :
        print ("total_hour : " + total_hour)

    # -10:00 (6)
    # -9:00 (5)
    # 1:00 (4)
    # 23:00 (5)
    # IF total_hour contains "-" 
    #   if length = 5 
    #       => place 0 between the - and number
    # else 
    #   if length = 4
    #       => place 0 before     
    if (total_hour.find("-") != -1) :
        if (len(total_hour) == 5) :
            total_hour = total_hour[1:]
            total_hour = "-0"+ total_hour
    else :
        if (len(total_hour) == 4) :
            total_hour = "0"+ total_hour

    if (debug) :
        print ("total_hour after manipulation : " + total_hour)


    # absences
    try :
        absences = tools.driver.find_element_by_xpath("/html/body/div[2]/div/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div[2]/ul/li/a/span[2]/span[1]/span[2]").text
        if (debug) :
            print ("absences before : " + absences)
        if (absences == 'Jour entier (7:24)') :
            absences = '7:24'
        elif (absences == '3:42') :
            absences = '3:42'
        elif (absences == u'Après-midi') :
            absences = '-3:42'
        else :
            absences = 'ERROR'

    except selenium.common.exceptions.NoSuchElementException:
        absences = 'Normal'
    if (debug) :
        print ("absences      : " + absences)


    informations = (
        in_1,
        out_1,
        in_2,
        out_2,
        total_hour,
        time_prested,
        absences
    )

    return informations

def createExcelFile(path, name_of_file, extension) :
    global workbook
    workbook = xlsxwriter.Workbook(excel_path + '/' + name_of_file + '.' + extension, {'date_1904': True}) 

def closeExcelFile() :
    global workbook
    workbook.close()


def writeInformationToExcel(date) :
    createFileInto(recoverInformation(), date )

def createFileInto(expenses, date ) :
    global workbook
    month = {
        '00': 'January',
        '01': 'January',
        '02': 'February',
        '03': 'March',
        '04': 'April',
        '05': 'May',
        '06': 'June',
        '07': 'July',
        '08': 'August',
        '09': 'September',
        '10': 'October',
        '11': 'November',
        '12': 'December'
    }
    column = {
        '0': 'A',
        '1': 'B',
        '2': 'C',
        '3': 'D',
        '4': 'E',
        '5': 'F',
        '6': 'G',
        '7': 'H',
        '8': 'I',
        '9': 'J',
        '10': 'K',
        '11': 'L',
        '12': 'M',
        '13': 'N',
        '14': 'O',
        '15': 'P',
        '16': 'Q',
        '17': 'R',
        '18': 'S',
        '19': 'T',
        '20': 'U',
        '21': 'V',
        '22': 'W',
        '23': 'X',
        '24': 'Y',
        '25': 'Z',
        '26': 'AA',
        '27': 'AB',
        '28': 'AC',
        '29': 'AD',
        '30': 'AE',
        '31': 'AF'
    }

    number_of_day_month = {
        'January': '31',
        'February': '28',
        'March': '31',
        'April': '30',
        'May': '31',
        'June': '30',
        'July': '31',
        'August': '31',
        'September': '30',
        'October': '31',
        'November': '30',
        'December': '31'
    }

    if (debug) :
        print ("date : " + date)
    month_from_date = date[5:-3]
    if (debug) :
        print ("month_from_date : " + month_from_date)
        print("equal to month : " + month[month_from_date])
    day_from_date = date[8:]
    if (debug) :
        print ("day_from_date : " + day_from_date)
        print("equal to colum : " + column[str(int(day_from_date))])
    letter = column[str(int(day_from_date))]
    #used to find the column before
    letter_before = column[str(int(day_from_date)-1)]
    if (debug) :
        print ("letter_before : " + letter_before)

    try:
        worksheet = workbook.add_worksheet(month[month_from_date])
    except xlsxwriter.exceptions.DuplicateWorksheetName:
        if (debug) :
            print("sheet already exist continue with the same")
        worksheet = workbook.get_worksheet_by_name((month[month_from_date]))
    
   

    # Start from the first cell. Rows and columns are zero indexed.
    row = 3
    col = int(day_from_date)

    # Write date
    worksheet.write(row - 1, col, date)

    # Create a format for the date or time.
    date_format = workbook.add_format({'num_format': '[hh]:mm', 'align': 'center'})
    for item in (expenses):
        worksheet.write(row, col, item, date_format)
        row += 1

    # Write a total using a formula.
    date_format2 = workbook.add_format({'num_format': '[hh]:mm', 'align': 'center'})
    date_format2.set_bg_color('gray')
    
    # Line 11
    worksheet.write(row, col, '='+letter+'5-'+letter+'4',date_format2 )
    
    # Line 12
    worksheet.write(row + 1, col, '='+letter+'7-'+letter+'6',date_format2 )
    
    # Line 13
    worksheet.write(row + 2, col, '=IF('+letter+'12+'+letter+'11>"09:00"*1,TIME(9,0,0)+'+letter+'16,'+letter+'12+'+letter+'11)',date_format2 )
    
    # Line 14
    worksheet.write(row + 3, col, '=IF(AND(IF('+letter+'11="00:00"*1,TRUE,FALSE),IF('+letter+'12="00:00"*1,TRUE,FALSE)),TIME(0,0,0),IF(OR(IF('+letter+'11="00:00"*1,TRUE,FALSE),IF('+letter+'12="00:00"*1,TRUE,FALSE)),TIME(3,42,0),TIME(7,24,0)))',date_format2 )

    # Line 15
    worksheet.write(row + 4, col, '=IF(OR(IF('+letter+'5="",TRUE,FALSE),IF('+letter+'6="",TRUE,FALSE)),TIME(0,0,0),'+letter+'6-'+letter+'5)',date_format2 )
    
    # Line 16
    worksheet.write(row + 5, col, '=IF(OR(IF('+letter+'11="00:00"*1,TRUE,FALSE),IF('+letter+'12="00:00"*1,TRUE,FALSE)),TIME(0,0,0),IF('+letter+'15<"00:30"*1,TIME(0,30,0)-('+letter+'15),TIME(0,0,0)))', date_format2 )
    
    # Line 17
    if (debug) :
        print (date)
    date_time_obj = datetime.datetime.strptime(date, '%Y-%m-%d')

    weekno = date_time_obj.weekday() 
    if weekno < 5 : 
        worksheet.write(row + 6, col, '=IF(AND('+letter+'10="7:24",ISBLANK('+letter+'4),ISBLANK('+letter+'6),'+letter+'9<>"0:00"),'+letter+'13-'+letter+'14-'+letter+'16-TIME(7,24,0),IF('+letter+'10="3:42", - (TIME(3,42,0)-'+letter+'11), IF('+letter+'10="-3:42", - (TIME(7,24,0)-'+letter+'11),'+letter+'13-'+letter+'14-'+letter+'16'+')))',date_format2 )
    else: 
        worksheet.write(row + 6, col, '00:00',date_format2)
    
    
    # Line 18
    worksheet.write(row + 7, col, '=IF(ISNUMBER(SEARCH("-",'+letter_before+'18)),IF('+letter+'17>=(TEXT(RIGHT('+letter_before+'18,LEN('+letter_before+'18)-FIND("-",'+letter_before+'18)),"[hh]:mm")),("-" & TEXT(RIGHT('+letter_before+'18,LEN('+letter_before+'18)-FIND("-",'+letter_before+'18))-'+letter+'17,"[hh]:mm")),TEXT('+letter+'17-RIGHT('+letter_before+'18,LEN('+letter_before+'18)-FIND("-",'+letter_before+'18)),"[hh]:mm")),TEXT('+letter_before+'18+'+letter+'17,"[hh]:mm"))', date_format2 ) 
    
    # Line 18
    green_format = workbook.add_format({'num_format': '[hh]:mm', 'align': 'center'})
    green_format.set_bg_color('green')
    worksheet.conditional_format(letter+'18', {'type': 'cell', 'criteria': 'equal to', 'value': '$'+letter+'$8', 'format': green_format})
    
    # Line 18
    red_format = workbook.add_format({'num_format': '[hh]:mm', 'align': 'center'})
    red_format.set_bg_color('red')
    worksheet.conditional_format(letter+'18', {'type': 'cell', 'criteria': 'not equal to', 'value': '$'+letter+'$8', 'format': red_format})

    # This place it's to revoverd the amount of hours from last month
    ## It's to have the opportunity to search into the list 00: Januray, 01: February, ...
    month_from_date_before = int(month_from_date) - 1
    if (month_from_date_before < 10) :
        month_before = "0" + str(int(month_from_date) - 1)
    else :
        month_before = str(int(month_from_date) - 1)
    
    # Line 18 column A
    if (debug) :
        print("equal to month : " + month[month_before])
        print("number_of_day_month : " + number_of_day_month[month[month_before]])
        print("number_of_day_month : " + column[number_of_day_month[month[month_before]]])
    # Exception when It's Januray we don't have the last entry of previous year.
    # For the moment take the hours in the first entry JANUARY B8
    if (month_from_date_before == 0) :
        worksheet.write(row + 7, 0, '=January!B8',date_format2 )
    else :
        worksheet.write(row + 7, 0, '='+month[month_before]+'!'+column[number_of_day_month[month[month_before]]]+'18',date_format2 )

    # Hide row 11 -> 17
    worksheet.set_row(10, None, None, {'hidden': True})
    worksheet.set_row(11, None, None, {'hidden': True})
    worksheet.set_row(12, None, None, {'hidden': True})
    worksheet.set_row(13, None, None, {'hidden': True})
    worksheet.set_row(14, None, None, {'hidden': True})
    worksheet.set_row(15, None, None, {'hidden': True})
    worksheet.set_row(16, None, None, {'hidden': True})
    
# Jacobs, L.J. (Laurent)
# https://nn.myprotime.eu/#/calendar/person/113787/month/daydetail/113787/2020-11-17?date=2020-01-01
# Thonon Pierre
# https://nn.myprotime.eu/#/calendar/person/113864/month/daydetail/113864/2020-11-18?date=2020-01-01


openBrowser()
connectToMyProtime()
enterCredentials()
goToMyCalendar()
createExcelFile(excel_path + '/', name_of_file, 'xlsx')

d1 = datetime.date(int(year_to_check), 1, 1)
if (begin_month_date != "") :
    d1 = datetime.date(int(year_to_check), int(begin_month_date), 1)

# d1 = datetime.date(2020, 6, 1)
today_year = datetime.datetime.today().strftime('%Y')
today_month = datetime.datetime.today().strftime('%m')
today_day = datetime.datetime.today().strftime('%d')
if (year_to_check != today_year) :
    if (end_month_date != "") :
        d2 = datetime.date(int(year_to_check), int(end_month_date), 31)
    else:
        d2 = datetime.date(int(year_to_check), 12, 31)
else :
    if (end_month_date != "") :
        d2 = datetime.date(int(today_year), int(end_month_date), int(today_day))
    else:
        d2 = datetime.date(int(today_year), int(today_month), int(today_day))
if (debug) :
    print (d2)

# d2 = datetime.date(2021, 1, 28)

days = [d1 + datetime.timedelta(days=x) for x in range((d2-d1).days + 1)]

for day in days:
    if (debug) :
        print(day.strftime('%Y-%m-%d'))
    goToMyCalendar_specific_date(day.strftime('%Y-%m-%d'))
    writeInformationToExcel(day.strftime('%Y-%m-%d'))

closeExcelFile()
closeBrowser()

# Close the program
tools.driver.quit()