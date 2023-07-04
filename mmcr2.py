import mmcr
from bs4 import BeautifulSoup as bs
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import time
from datetime import date


mmcr2 = "https://mmcr-amap.i.daimler.com"


def login(driver,user,password):
    """Helper function for logging into MMCR 2.0"""

    #get mmcr directly, avoid netstar as mmcr directly requires no 2FA
    driver.get(mmcr2)

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
    
    #closeButton = WebDriverWait(driver,45).until(ec.presence_of_element_located((By.CLASS_NAME,"cookiebanner-close")))
    #action = webdriver.common.action_chains.ActionChains(driver)
    #action.move_to_element_with_offset(closeButton, 0, 0)
    #action.click()
    #action.perform()

    return driver


def checkServices(user=None,password=None):
    """Function for scraping service status off of MMCR 2.0, used in conjunction with Elead module's output table."""

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
                    #sc-gzrROc jzOepy sc-fmPOXC gqHAVM active
                    accountLink = WebDriverWait(driver,15).until(ec.presence_of_element_located((By.ID,"sc-gzrROc jzOepy sc-fmPOXC gqHAVM active")))
                    accountLink.click()
                except:
                    pass
                
                #search email
                usernameElement = WebDriverWait(driver,10).until(ec.presence_of_element_located((By.ID,"ME_ID_DASHBOARD")))
                usernameElement.send_keys(e)
                searchButton = driver.find_element(By.ID,"sc-laZRCg gaooqp")
                searchButton.click()

                try:
                    #wait for vehicles tab to popup (if found this means we found a profile)
                    vehicleTable = WebDriverWait(driver,4).until(ec.presence_of_element_located((By.ID,"sc-hZNxer dcaJSz")))
                    vehicleTableSoup = bs(vehicleTable.get_attribute("innerHTML"),'html.parser')
                    tbody = vehicleTableSoup.find("tbody")
                    tableRow = 0

                    for tr in tbody.findAll("tr"):
                        pass

                except:
                    #if no profile is found exception is thrown and we handle according to what went wrong
                    #src = driver.page_source
                    #soup = bs(src,'html.parser')
                    #if soup.find("div",class_="modal modal-alert-mini fade in show"):
                    #    #profile missing details, in this case that is their profile so we can stop searching for their other possible emails
                    #    closePopup = driver.find_element(By.ID, "mmcr:confirmMissingProfileFields:close:command")
                    #    closePopup.click()
                    #    statusTemp = "MMCR profile requires update"
                    #    break
                    try:
                        closePopup = WebDriverWait(driver,1).until(ec.presence_of_element_located((By.CLASS_NAME,"sc-laZRCg durlsR")))
                        closePopup.click()
                        continue
                    except:
                        continue