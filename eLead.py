from bs4 import BeautifulSoup as bs
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import time
import datetime

loginPage = "https://www.eleadcrm.com/evo2/fresh/login.asp"
Appointments = "https://www.eleadcrm.com/evo2/fresh/eLead-V45/elead_track/servicesched/serviceappointments.aspx"
iframe = "https://www.eleadcrm.com/evo2/fresh/eLead-V45/elead_track/search/searchresults.asp?Go=2&searchexternal=&corpStore=False&st=0&lIUID=&etitle=&lDID=&PID=&origStrDo=#"
dayOneHTML = None
dayTwoHTML = None
USWorkHoldays = ["01-01", "07-04", "09-09", "11-11", "02-22", "12-25", "11-25"]


def HTMLtoExcel(HTML,date,indexstart=0):
    """Parses table HTML and filters for only appointments of model years 2019-2022 then stores in excel"""

    list = []

    soup = bs(HTML, 'html.parser').find_all('tbody')[0]
    
    for row in soup.findAll("tr"):
        car = row.find("td", class_="vehicle")
        year = int(car.text[0:4])
        if(2019<=year<2023):
            VINdex = car.text.find("VIN")
            vin = car.text[VINdex+4:]
            model = car.text[0:VINdex]
            time = row.find("td", class_="time").text
            rep = row.find("td", class_="advisor").text
            name = row.find("td", class_="customer").text
            list.append([date,time,name,model,vin,rep])
            
        else:
            pass
    #create dataframe out of filtered appointments and send to excel

    index = range(indexstart,indexstart+len(list))

    df = pd.DataFrame(list,columns=['Date', 'Time', 'Customer Name', 'Vehicle', 'VIN', 'Service Rep'], index=index)

    return df


def getSchedule(driver,email=None,password=None):
    """Gets service appointment table from Eleads and filters out any appointments that
     the vehicle is not a 2019-2022 (only 2019-2022 can possibly have expired services)"""
    
    #figure out the two next business days of the week to grab
    dayOne,dayTwo = getDate()
    days = [dayOne.isoweekday(),dayTwo.isoweekday()]
    
    #login process
    driver.get(loginPage)
    SSOlogin = WebDriverWait(driver,20).until(ec.presence_of_element_located((By.ID,"btnSingleSignOnSignUp")))
    SSOlogin.click()
    
    #send and submit credentials
    emailInput = WebDriverWait(driver,20).until(ec.presence_of_element_located((By.ID,"okta-signin-username")))
    passInput = driver.find_element("id","okta-signin-password")
    submit = driver.find_element("id","okta-signin-submit")
    emailInput.send_keys(email)
    passInput.send_keys(password)
    submit.click()

    #request sms 2FA verification and wait until Eleads main menu loads to ensure auth cookies are set
    smsRequestButton = WebDriverWait(driver,20).until(ec.presence_of_element_located((By.XPATH,"/html/body/div[2]/div[2]/main/div[2]/div/div/form/div[1]/div[2]/a")))
    smsRequestButton.click()
    WebDriverWait(driver,60).until(ec.presence_of_element_located((By.ID,"divMainMenu")))

    #go to service appointment page (skip loading border html and just grab content,
    #hence why i use get instead of frontend... plus makes it slightly easier to grab table elements)
    driver.get(Appointments)
    time.sleep(2)

    #if the first day is 1 (monday), we must click the next week arrow, as that means we are grabbing next weeks tables
    if days[0] == 1:
        nextElement = driver.find_element("id","nextSegment")
        nextElement.click()
        time.sleep(1)

    #navigate to appropriate day and grab the tables html to parse
    dayButton = getDayButton(driver,days[0])
    dayButton.click()
    time.sleep(1)
    appointmentTable = driver.find_element("id","appointmentContainer")
    dayOneHTML = appointmentTable.get_attribute("innerHTML")

    #if the second day is 1 (monday), we must click the next week arrow, as that means we are grabbing next weeks tables
    if days[1] == 1:
        nextElement = driver.find_element("id","nextSegment")
        nextElement.click()
        time.sleep(1)

    #navigate to appropriate day and grab the tables html to parse
    dayButton = getDayButton(driver,days[1])
    dayButton.click()
    time.sleep(1)
    appointmentTable = driver.find_element("id","appointmentContainer")
    dayTwoHTML = appointmentTable.get_attribute("innerHTML")

    #parse and filter table, then store in excel sheet
    df1 = HTMLtoExcel(dayOneHTML, str(dayOne))
    df1s = df1.shape
    df2 = HTMLtoExcel(dayTwoHTML, str(dayTwo),indexstart=df1s[0])
    df = pd.concat([df1,df2],verify_integrity=True)

    return df


def getDayButton(ChromeDriver,day):
    """Gets day element and returns it."""

    if day == 1:
        element = ChromeDriver.find_element("id","Mon")
    elif day == 2:
        element = ChromeDriver.find_element("id","Tue")
    elif day == 3:
        element = ChromeDriver.find_element("id","Wed")
    elif day == 4:
        element = ChromeDriver.find_element("id","Thu")
    elif day == 5:
        element = ChromeDriver.find_element("id","Fri")
    return element


def getDate():
    """Grabs next business date according to offset (if offset is 2 and today is thursday, returns monday's date)"""
    today = datetime.date.today()
    offset = 1
    
    while(True):
        if offset == 1 and today.isoweekday() == 5:
            dateOne = today + datetime.timedelta(days=3)
        elif offset == 1 and today.isoweekday() == 6:
            dateOne = today + datetime.timedelta(days=2)
        elif offset == 1 and today.isoweekday() == 7:
            dateOne = today + datetime.timedelta(days=1)
        else:
            dateOne = today + datetime.timedelta(days=offset)
        HolidayCheck = dateOne.strftime('%m-%d')
        
        if USWorkHoldays.count(HolidayCheck):
            today = dateOne
        else:
            break

    today = dateOne
    while(True):
        if offset == 1 and today.isoweekday() == 5:
            dateTwo = today + datetime.timedelta(days=3)
        elif offset == 1 and today.isoweekday() == 6:
            dateTwo = today + datetime.timedelta(days=2)
        elif offset == 1 and today.isoweekday() == 7:
            dateTwo = today + datetime.timedelta(days=1)
        else:
            dateTwo = today + datetime.timedelta(days=offset)
        HolidayCheck = dateTwo.strftime('%m-%d')
        
        if USWorkHoldays.count(HolidayCheck):
            today = dateTwo
        else:
            break

    return dateOne,dateTwo