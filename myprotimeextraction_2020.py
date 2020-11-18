from Tools import tools_v000 as tools
import os
from os.path import dirname
import selenium
from selenium.webdriver.common.keys import Keys
import xlsxwriter
import datetime


# -24 for the name of this project myProtimeExtraction_2020
save_path = dirname(__file__)[ : -24]
propertiesFolder_path = save_path + "Properties"

# Properties
email_text = tools.readProperty(propertiesFolder_path, 'myProtimeExtraction_2020', 'Email=')
password_text = tools.readProperty(propertiesFolder_path, 'myProtimeExtraction_2020', 'Password=')
excel_path = tools.readProperty(propertiesFolder_path, 'myProtimeExtraction_2020', 'excel_path=')

# static variable
workbook = ''
worksheet = ''

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
    tools.waitLoadingPageByID2(30, 'day-program')
    try :
        in_1 = tools.driver.find_element_by_xpath("/html/body/div[2]/div/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div/div[3]/div[2]/div/ul/li[1]/span[1]/a").text
    except selenium.common.exceptions.NoSuchElementException:
        in_1 = '07:00'
    print ("in_1       : " + in_1)
    try :
        out_1 = tools.driver.find_element_by_xpath("/html/body/div[2]/div/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div/div[3]/div[2]/div/ul/li[2]/span[1]/a").text
    except selenium.common.exceptions.NoSuchElementException:
        out_1 = '12:00'
    print ("out_1      : " + out_1)
    try :
        in_2 =  tools.driver.find_element_by_xpath("/html/body/div[2]/div/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div/div[3]/div[2]/div/ul/li[3]/span[1]/a").text
    except selenium.common.exceptions.NoSuchElementException:
        in_2 = '12:30'
    print ("in_2       : " + in_2)
    try :
        out_2 = tools.driver.find_element_by_xpath("/html/body/div[2]/div/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div/div[3]/div[2]/div/ul/li[4]/span[1]/a").text
    except selenium.common.exceptions.NoSuchElementException:
        out_2 = '14:54'
    print ("out_2      : " + out_2)
    try :                                                
        total_hour = tools.driver.find_element_by_xpath("/html/body/div[2]/div/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div/div[3]/div[1]/ul/li[5]/span/span[2]").text
    except selenium.common.exceptions.NoSuchElementException:
        total_hour = tools.driver.find_element_by_xpath("/html/body/div[2]/div/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div/div[3]/div[1]/ul/li[4]/span/span[2]").text
    print ("total_hour : " + total_hour)

    informations = (
        in_1,
        out_1,
        in_2,
        out_2,
        total_hour
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

    print ("date : " + date)
    month_from_date = date[5:-3]
    print ("month_from_date : " + month_from_date)
    print("equal to month : " + month[month_from_date])
    day_from_date = date[8:]
    print ("day_from_date : " + day_from_date)
    print("equal to colum : " + column[str(int(day_from_date))])
    letter = column[str(int(day_from_date))]
    #used to find the column before
    letter_before = column[str(int(day_from_date)-1)]
    print ("letter_before : " + letter_before)

    try:
        worksheet = workbook.add_worksheet(month[month_from_date])
    except xlsxwriter.exceptions.DuplicateWorksheetName:
        print("sheet already exist continue with the same")
        worksheet = workbook.get_worksheet_by_name((month[month_from_date]))
    
   

    # Start from the first cell. Rows and columns are zero indexed.
    row = 3
    col = int(day_from_date)

    # Write date
    worksheet.write(row - 1, col, date)

    # Create a format for the date or time.
    date_format = workbook.add_format({'num_format': 'hh:mm', 'align': 'center'})
    for item in (expenses):
        worksheet.write(row, col, item, date_format)
        row += 1

    # Write a total using a formula.
    date_format2 = workbook.add_format({'num_format': 'hh:mm', 'align': 'center'})
    date_format2.set_bg_color('gray')
    worksheet.write(row, col, '='+letter+'5-'+letter+'4',date_format2 )
    worksheet.write(row + 1, col, '='+letter+'7-'+letter+'6',date_format2 )
    worksheet.write(row + 2, col, '='+letter+'10+'+letter+'9',date_format2 )
    worksheet.write(row + 3, col, '07:24',date_format2 )
    worksheet.write(row + 4, col, '='+letter+'6-'+letter+'5',date_format2 )
    worksheet.write(row + 5, col, '=IF('+letter+'13<"00:30"*1,TIME(0,30,0)-('+letter+'6-'+letter+'5),TIME(0,0,0))', date_format2 )
    worksheet.write(row + 6, col, '='+letter+'11-'+letter+'12-'+letter+'14',date_format2 )
    worksheet.write(row + 7, col, '='+letter_before+'16+'+letter+'15',date_format2 )
    
    
    worksheet.write(row + 7, 0, '07:24',date_format2 )


# Jacobs, L.J. (Laurent)
# https://nn.myprotime.eu/#/calendar/person/113787/month/daydetail/113787/2020-11-17?date=2020-01-01
# Thonon Pierre
# https://nn.myprotime.eu/#/calendar/person/113864/month/daydetail/113864/2020-11-18?date=2020-01-01


openBrowser()
connectToMyProtime()
enterCredentials()
goToMyCalendar()
createExcelFile(excel_path + '/', "test", 'xlsx')

d1 = datetime.date(2020, 1, 1)
today_year = datetime.datetime.today().strftime('%Y')
today_month = datetime.datetime.today().strftime('%m')
today_day = datetime.datetime.today().strftime('%d')
d2 = datetime.date(int(today_year), int(today_month), int(today_day)-1)

print (d2)

days = [d1 + datetime.timedelta(days=x) for x in range((d2-d1).days + 1)]

for day in days:
    print(day.strftime('%Y-%m-%d'))
    goToMyCalendar_specific_date(day.strftime('%Y-%m-%d'))
    writeInformationToExcel(day.strftime('%Y-%m-%d'))

closeExcelFile()
closeBrowser()

# goToMyCalendar_specific_date('2020-10-01')
# writeInformationToExcel('2020-10-01')

# goToMyCalendar_specific_date('2020-10-02')
# writeInformationToExcel('2020-10-02')

# goToMyCalendar_specific_date('2020-10-03')
# writeInformationToExcel('2020-10-03')

# goToMyCalendar_specific_date('2020-10-05')
# writeInformationToExcel('2020-10-05')

# goToMyCalendar_specific_date('2020-11-17')
# writeInformationToExcel('2020-11-17')





# createFileInto(excel_path + '/', "test", 'xlsx', '')
