from bs4 import BeautifulSoup as bs
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import codecs
import time
import datetime


loginPage = "https://www.eleadcrm.com/evo2/fresh/login.asp"
email = ""
password = ""
Appointments = "https://www.eleadcrm.com/evo2/fresh/eLead-V45/elead_track/servicesched/serviceappointments.aspx"
iframe = "https://www.eleadcrm.com/evo2/fresh/eLead-V45/elead_track/search/searchresults.asp?Go=2&searchexternal=&corpStore=False&st=0&lIUID=&etitle=&lDID=&PID=&origStrDo=#"
dayOneHTML = None
dayTwoHTML = None


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


def getSchedule(driver,email=email,password=password):
    """Gets service appointment table from Eleads and filters out any appointments that
     the vehicle is not a 2019-2022 (only 2019-2022 can possibly have expired services)"""
    
    #figure out the two next business days of the week to grab
    days = getDays()
    
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
    df1 = HTMLtoExcel(dayOneHTML, str(getDate(1)))
    df1s = df1.shape
    df2 = HTMLtoExcel(dayTwoHTML, str(getDate(2)),indexstart=df1s[0])
    df = pd.concat([df1,df2],verify_integrity=True)

    return df,driver
    

def getDays():
    """determine next two business days (skips sat and sun)"""

    today = datetime.date.today()
    if today.isoweekday() == 4:
        dayOne = 5
        dayTwo = 1
    elif today.isoweekday() == 5:
        dayOne = 1
        dayTwo = 2
    else:
        dayOne = today.isoweekday() + 1
        dayTwo = today.isoweekday() + 2
    days = [dayOne,dayTwo]
    return days


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


def getDate(offset):
    """Grabs next business date according to offset (if offset is 2 and today is thursday, returns monday's date)"""
    today = datetime.date.today()
    if offset == 1 and today.isoweekday() == 5:
        date = today + datetime.timedelta(days=3)
    elif offset == 2 and (today.isoweekday() == 5 or today.isoweekday() == 4):
        date = today + datetime.timedelta(days=4)
    else:
        date = today + datetime.timedelta(days=offset)
    return date


def pullEmails(driver,df):
    """Navigates Eleads and searches for all of the names in the service appointment dataframe, grabs all emails associated with the account if the 
    correct VIN is linked to the Eleads account."""

    #read in excel data
    #df = pd.read_excel(r'out.xlsx')
    #go back from service appointment so that searchbar is available
    driver.back()

    #turn customer name and VIN column's into a list
    nameList = df['Customer Name'].to_list()
    vinList = df['VIN'].to_list()
    #initialize email list that will be put into dataframe
    emailList = []

    #enumerate so we can reference VIN list aswell
    for i,name in enumerate(nameList):
        
        if len(driver.window_handles)>1:
                driver.switch_to.window(driver.window_handles[1])
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
        
        #wait until searchbar is available, send keys, then search
        searchBar = WebDriverWait(driver,20).until(ec.presence_of_element_located((By.ID,"txtQuickSearch")))
        searchBar.send_keys(name)
        search = driver.find_element("id","tdSearchImage")
        search.click()

        #grab iframe content directly
        driver.get(iframe)
        
        #grab table element
        tableElement = driver.find_element(By.XPATH,"/html/body/table/tbody/tr[2]/td/table/tbody")
        row = tableElement.find_elements(By.TAG_NAME,"tr")
       
        emails = ""
        #iterate through first four rows of search results (cap at 4 to save time and resources)
        for j in range(0,3):
            tempEmails=""

            #attempt to grab table data from current iterations row, else break loop
            try:
                td = row[j].find_elements(By.TAG_NAME,"td")
            except:
                break
            #close excess windows (error prevention for handle switching)
            if len(driver.window_handles)>1:
                driver.switch_to.window(driver.window_handles[1])
                driver.close()
                driver.switch_to.window(driver.window_handles[0])

            #only check users with 100% match of search info
            if td[1].text == "100%":
                #open profile
                infoButton = row[j].find_element("id","moreInfo"+str(j))
                infoButton.click()
                
                #profile is popup window so switch control to said window 
                driver.switch_to.window(driver.window_handles[1])

                #allow page to load, but bypass iframe loads
                WebDriverWait(driver,60).until(ec.presence_of_element_located((By.ID,"trFrame")))
                
                #the way the page is structured I couldn't find a clean way to grab emails off of frontend (besides secondary emails as they had ID on TR)
                #so insted we pull source code and split the data off of the script setting data
                try:
                    src = driver.page_source
                    soup = bs(src, 'html.parser')
                    scripts = soup.find_all('script',type="text/javascript")
                    scripts=scripts[52].text
                    split = scripts.split("var g_data = {")[1]
                    g_data = split.split("}",1)[0]
                except:
                    #error handling incase we grab source code too early, waits 5s then tries again
                    time.sleep(5)
                    src = driver.page_source
                    soup = bs(src, 'html.parser')
                    scripts = soup.find_all('script',type="text/javascript")
                    scripts=scripts[52].text
                    split = scripts.split("var g_data = {")[1]
                    g_data = split.split("}",1)[0]
                primaryEmail = g_data.split("PrimaryEmail:")[1][1:].split("\"")[0]
                #add primary email to temp emails
                tempEmails += primaryEmail

                #attempt to find other emails, add those to tempEmails using / as a delimiter
                try:
                    otherEmailsRow = soup.find("tr",id="CustomerInfoPanel_OtherEmailsRows")
                    otherEmailsTable = otherEmailsRow.find("table")
                    for tr in otherEmailsTable.find_all("tr"):
                        if tr.find("td").text.find('@')!=-1:
                            tempEmails = tempEmails+'/'+tr.find("td").text
                except(AttributeError,NameError):
                    pass

                #build URL for vehicle information to check VIN's and send get request
                url = driver.current_url
                ipid = url.split("lPID=")[1].split('&')[0]
                vinURL = "https://www.eleadcrm.com/evo2/fresh/eLead-V45/elead_track/NewProspects/custvehicles.aspx?tab=liVehicles&lICID=5150&lPID=" + ipid + "&lIUID=3634215"
                driver.get(vinURL)
                WebDriverWait(driver,20).until(ec.presence_of_element_located((By.ID,"div_purchased")))

                #if vin of vehicle coming in for service is linked to his eleads profile then add our temp emails to email to be added to list
                try:
                    soup = bs(driver.page_source,'html.parser')
                    purchasedTable = soup.find(id="gvPurchased")
                    rows = purchasedTable.findAll("tr")
                    for tr in rows:
                        if str(tr.find("td").text).find(str(vinList[i]))!=-1 and len(str(vinList[i]))==17:
                            emails += tempEmails
                except(AttributeError,NameError):
                    pass
                
                #close popup window and switch control to main window
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
        # if we couldn't find an email, or an account with the VIN linked set email as "?"        
        if emails=="":
            emails = '?'
        #append email to list
        emailList.append(emails)
        #go to previous page to use search bar
        driver.back()
    #add email list to dataframe
    df['Email(s)'] = emailList
    
    #write to excel
    with pd.ExcelWriter("out.xlsx",mode="w") as writer:
        df.to_excel(excel_writer = writer,index=False)

    #close selenium instances
    driver.quit()


def elead(email,password):
    driver = webdriver.Chrome()
    df,driver = getSchedule(driver,email=email,password=password)
    pullEmails(driver,df)