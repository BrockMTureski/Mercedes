from bs4 import BeautifulSoup as bs
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import time
from datetime import date
import requests

netStar = "https://netstar.i.mercedes-benz.com/"
mmcr = "https://mmcr-amap.i.daimler.com/MMCR/"

def login(driver,user,password):
    """Helper function for logging into MMCR."""

    #get mmcr directly, avoid netstar as mmcr directly requires no 2FA
    driver.get(mmcr)

    #get username text box and send username
    usernameElement = WebDriverWait(driver,10).until(ec.presence_of_element_located((By.ID,"userid")))
    usernameElement.send_keys(user)
    #submit username
    nextButton = driver.find_element(By.ID,"next-btn")
    nextButton.click()

    #get password text box and send password
    passwordElement = WebDriverWait(driver,10).until(ec.presence_of_element_located((By.ID,"password")))
    passwordElement.send_keys(password)
    #submit password
    loginButton = driver.find_element(By.ID,"loginSubmitButton")
    loginButton.click()

    #wait for cookie banner to popup (website loaded at that point), close banner once element available
    closeButton = WebDriverWait(driver,45).until(ec.presence_of_element_located((By.CLASS_NAME,"cookiebanner-close")))
    action = webdriver.common.action_chains.ActionChains(driver)
    action.move_to_element_with_offset(closeButton, 0, 0)
    action.click()
    action.perform()

    return driver


def checkServices(user=None,password=None):
    """Function for scraping service status off of MMCR, used in conjunction with Elead module's output table."""

    driver = webdriver.Chrome()

    #read in output from elead module, then split email and vin to list and initialize status and expiry lists
    df = pd.read_excel(r'out.xlsx')
    emailList = df['Email(s)'].to_list()
    vinList = df['VIN'].to_list()
    statusList = []
    expiryList = []
    
    #login to mmcr
    driver = login(driver,user,password)

    #loop through each email to attempt services
    for i,email in enumerate(emailList):
        statusTemp = ""
        expiry = ""

        #split email into all possible emails, each seperated by delimiter
        localEmail=email.split('/')

        #if email is displayed as "?" dont bother with inner loop as eleads module couldn't find
        #any emails
        if "?" not in email[0]: 
            #loop through all of the possible emails in each profile
            for e in localEmail:
                #if a status was found then break loop as rest of possible emails dont need to be checked
                if statusTemp!="":
                    break
                #go to account search screen
                try:
                    accountLink = WebDriverWait(driver,15).until(ec.presence_of_element_located((By.ID,"mmcr:primaryMenu:accounts:link")))
                    accountLink.click()
                except:
                    #error handling (MMCR is very unreliable so basically any time a request is made we error handle like this)
                    #just takes us back to account and goes to next profile
                    driver.back()
                    driver.refresh()
                    time.sleep(12)
                    backToMMCR = driver.find_element(By.XPATH,"/html/body/form/div[1]/main/div/div/div/div/div[3]/div/a")
                    backToMMCR.click()
                    statusTemp = "Error Account Info"
                    break
                
                #search email in accounts tab
                usernameElement = WebDriverWait(driver,10).until(ec.presence_of_element_located((By.ID,"mmcr:accountSearch:accountIdentifier:input")))
                usernameElement.send_keys(e)
                searchButton = driver.find_element(By.ID,"mmcr:accountSearch:searchById:command")
                searchButton.click()
                
                try:
                    #wait for vehicles tab to popup (if found this means we found a profile)
                    vehicleTab = WebDriverWait(driver,6).until(ec.presence_of_element_located((By.ID,"mmcr:accountOverview:accountVehicles:tab")))
                    vehicleTab.click()
                except:
                    #if no profile is found exception is thrown and we handle according to what went wrong
                    src = driver.page_source
                    soup = bs(src,'html.parser')
                    if soup.find("div",class_="modal modal-alert-mini fade in show"):
                        #profile missing details, in this case that is their profile so we can stop searching for their other possible emails
                        closePopup = driver.find_element(By.ID, "mmcr:confirmMissingProfileFields:close:command")
                        closePopup.click()
                        statusTemp = "MMCR profile requires update"
                        break

                    elif soup.find("div",class_="modal-content"):
                        #Incorrect email format
                        closePopup = driver.find_element(By.ID,"mmcr:message:close:command")
                        closePopup.click()
                        continue

                    else:
                        #can't find profile condition, go to next email
                        continue
                
                try:
                    #attempt to find table listing vehicles
                    vehicles =  WebDriverWait(driver,10).until(ec.presence_of_element_located((By.ID,"mmcr:accountOverview:assignedVehicles:tbody_element")))
                except: 
                    break

                #pass to beautiful soup
                vehicleTableSoup = bs(vehicles.get_attribute("innerHTML"),'html.parser')
                tableRow = 0
                found = False
                
                #search through vehicle table to figure out which vehicle matches the VIN
                for tr in vehicleTableSoup.findAll("tr"):
                    td = tr.findAll("td")
                    if len(td)<2:
                        break
                    
                    check = td[2].text.split()
                    if len(check)==1:
                        check = check[0][-17:]
                    else:
                        check = check[1][-17:]
                    if check == vinList[i]:
                        found = True
                        break
                    else:
                        tableRow += 1

                #if no vehicle on profile matches the VIN check next email
                if not found:
                    continue
                else:
                    #if vehicle with matching vin is found click on vehicle to check services
                    vehicleID = "mmcr:accountOverview:assignedVehicles:"+str(tableRow)+":select:command"
                    vehicleElement = driver.find_element(By.ID, vehicleID)
                    vehicleElement.click()

                    try:
                        #click on services tab
                        servicesTab = WebDriverWait(driver,5).until(ec.presence_of_element_located((By.ID,"mmcr:vehicleOverview:vehicleServices:tab")))
                        servicesTab.click()
                    except:
                        #error handling
                        time.sleep(5)
                        driver.back()
                        driver.refresh()
                        time.sleep(12)
                        backToMMCR = driver.find_element(By.XPATH,"/html/body/form/div[1]/main/div/div/div/div/div[3]/div/a")
                        backToMMCR.click()
                        statusTemp = "Error Vehicle Info"
                        break
                    
                    #wait until page is loaded
                    WebDriverWait(driver,10).until(ec.presence_of_element_located((By.ID,"mmcr:vehicleOverview:services")))
                    #grab page source and pass to BS4 for easier parsing
                    pagesrc = driver.page_source
                    serviceTableSoup = bs(pagesrc, 'html.parser').find("tbody",id="mmcr:vehicleOverview:services:tbody_element")

                    #iterate through rows until we find "Remote Door Lock & Unlock" row
                    for tr in serviceTableSoup.findAll("tr"):
                        td = tr.findAll("td")
                       
                        if td[1].text == "Remote Door Lock & Unlock":
                            try:
                                #status temp will be either "Activated" or "Deactivated"
                                statusTemp = td[4].find("button")['data-content']
                                
                                #if activated grab when it expires
                                if statusTemp == "Activated":
                                    expiry = td[3].text[-10:]

                            except(NameError,AttributeError):
                                pass
                #if status found no need to check rest of emails for this client
                if statusTemp == "Deactivated" or "Activated":
                    break
        #change empty status's to a '?' just to indicate to look into this client by hand
        if statusTemp == "":
            statusTemp = "?"
        #append status to list
        statusList.append(statusTemp)
        #append expiry to list, place "N/A" if status is not activated
        if expiry!="":
            expiryList.append(expiry)
        else:
            expiryList.append("N/A")

    #add status list and expiry list to dataframe
    df['Status'] = statusList
    df['Expiry'] = daysBetween(expiryList)

    #send find dataframe to excel
    with pd.ExcelWriter("status1.xlsx",mode="w") as writer:
        df.to_excel(excel_writer = writer,index=False)


def daysBetween(expiry):
    """Function takes in a list of expiry dates and determines how many days between now and expiry,
      then adds how many days until expiry next to expiry date in the list."""
    
    today = date.today()

    for i,expiryDate in enumerate(expiry):
        if str(expiryDate)=="N/A":
            continue
        else:
            split = str(expiryDate)[0:10].split('/')
            year = int(split[2])
            month = int(split[0])
            day = int(split[1])
            d = date(year,month,day)
            dayNum = (d - today).days
            expiry[i] = str(expiry[i])[0:10] +' (' + str(dayNum) + ' days)'
    return expiry