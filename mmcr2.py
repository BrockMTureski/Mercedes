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

    driver.maximize_window()

    return driver


def checkServices(user=None,password=None):
    """Function for scraping service status off of MMCR 2.0, used in conjunction with Elead module's output table."""

    driver = webdriver.Chrome()

    #read in output from elead module, then split email and vin to list and initialize status and expiry lists
    df = pd.read_excel(r'out.xlsx')
    emailList = df['Email(s)'].to_list()
    vinList = df['VIN'].to_list()
    nameList = df['Customer Name'].to_list()
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
                    accountLink = WebDriverWait(driver,4).until(ec.presence_of_element_located((By.XPATH,"/html/body/div[1]/header/div/div[2]/div[1]/div/a[1]")))
                    accountLink.click()
                except:
                    pass
                
                #search email
                usernameElement = WebDriverWait(driver,10).until(ec.presence_of_element_located((By.ID,"ME_ID_DASHBOARD")))
                usernameElement.send_keys(e)
                try:
                    searchButton = WebDriverWait(driver,1).until(ec.presence_of_element_located((By.XPATH,"/html/body/div[1]/div[2]/div/div/div[2]/div/div[2]/div/div[2]/div/form/div[2]/div[2]/button")))
                    searchButton.click()
                except:
                    print("Incorrect email formatting")


                #wait for vehicles tab to popup (if found this means we found a profile)
                try:
                    time.sleep(4)
                    check = driver.current_url
                    if check != "https://mmcr-amap.i.daimler.com/account":
                        raise Exception("Account Not Found")
                    
                    vehicleTable = WebDriverWait(driver,10).until(ec.presence_of_element_located((By.XPATH,"/html/body/div[1]/div/div[2]/div/div/div[2]/div[1]/div[2]/table")))
                    vehicleTableSoup = bs(vehicleTable.get_attribute("innerHTML"),'html.parser')
                    tbody = vehicleTableSoup.find("tbody")
                    row = None

                    for num,tr in enumerate(tbody.findAll("tr")):
                        if tr.findAll("th")[1].text == vinList[i]:
                            row = num+1

                    if row is not None:
                        xpath = "/html/body/div[1]/div/div[2]/div/div/div[2]/div[1]/div[2]/table/tbody/tr[" + str(row) + "]/th[7]/a"
                        vehicleButton = driver.find_element(By.XPATH, xpath)
                        vehicleButton.click()
                        
                        serviceTable = WebDriverWait(driver,8).until(ec.presence_of_element_located((By.XPATH,"/html/body/div[1]/div/div[2]/div/div[4]/div/div[2]/div/div[2]/div[1]/table/tbody")))
                        serviceTableSoup = bs(serviceTable.get_attribute("innerHTML"),'html.parser')
                        serviceTableRows = serviceTableSoup.find_all("tr")
                        for row in serviceTableRows:
                            data = row.find_all("th")
                            if data[1].text == "Remote Door Lock & Unlock":
                                if len(data[3].text) != 0:
                                    if data[2].text == " Information required":
                                        statusTemp = "Actived (Profile Requires Attention)"
                                    else:
                                        statusTemp = "Activated"
                                    expiry = data[3].text[-10:]     
                                else:
                                    statusTemp = "Deactivated"
                                break

                except Exception as e:
                    try:
                        closePopup = WebDriverWait(driver,1).until(ec.presence_of_element_located((By.XPATH,"/html/body/div[6]/div/div/footer/div[1]/button")))
                        closePopup.click()
                        statusTemp = "MMCR Profile Requires Update"
                    except:
                        continue

        if statusTemp=="":
            statusTemp = "?"
        if expiry == "":
            expiry = "N/A"
        statusList.append(statusTemp)
        expiryList.append(expiry)
    
    for index,status in enumerate(statusList):
        if status == '?':

            accountLink = WebDriverWait(driver,4).until(ec.presence_of_element_located((By.XPATH,"/html/body/div[1]/header/div/div[2]/div[1]/div/a[1]")))
            accountLink.click()
            
            vinSearch = WebDriverWait(driver,4).until(ec.presence_of_element_located((By.XPATH,"/html/body/div[1]/div[2]/div/div/div[2]/div/div[2]/div/div[1]/ol/button[3]")))
            vinSearch.click()
            vinInput = WebDriverWait(driver,4).until(ec.presence_of_element_located((By.XPATH,"/html/body/div[1]/div[2]/div/div/div[2]/div/div[2]/div/div[2]/div/div/form/div[1]/div/div/div/input")))
            vinInput.send_keys(vinList[index])
            fullName = nameList[index].split()
            if len(fullName)<2:
                statusList[index] = "Error: Name Length"
                continue
            try:
                time.sleep(3)
                firstNameInput = WebDriverWait(driver,2).until(ec.presence_of_element_located((By.XPATH,"/html/body/div[1]/div[2]/div/div/div[2]/div/div[2]/div/div[2]/div/div/form/div[2]/div[1]/div/div/input")))
                lastNameInput = WebDriverWait(driver,2).until(ec.presence_of_element_located((By.XPATH,"/html/body/div[1]/div[2]/div/div/div[2]/div/div[2]/div/div[2]/div/div/form/div[2]/div[2]/div/div/input")))
                firstNameInput.send_keys(fullName[0])
                lastNameInput.send_keys(fullName[1])
                
            except:
                statusList[index] = "Not Paired"
                continue
            search = WebDriverWait(driver,2).until(ec.presence_of_element_located((By.XPATH,"/html/body/div[1]/div[2]/div/div/div[2]/div/div[2]/div/div[2]/div/div/form/div[3]/button")))
            search.click()
            try:
                serviceTable = WebDriverWait(driver,7).until(ec.presence_of_element_located((By.XPATH,"/html/body/div[1]/div/div[2]/div/div[4]/div/div[2]/div/div[2]/div[1]/table/tbody")))
                serviceTableSoup = bs(serviceTable.get_attribute("innerHTML"),'html.parser')
                serviceTableRows = serviceTableSoup.find_all("tr")
                for row in serviceTableRows:
                    data = row.find_all("th")
                    if data[1].text == "Remote Door Lock & Unlock":
                        if len(data[3].text) != 0:
                            if data[2].text == " Information required":
                                statusList[index] = "Actived (Profile Requires Attention)"
                            else:
                                statusList[index] = "Activated"
                            expiryList[index] = data[3].text[-10:]     
                        else:
                            statusList[index] = "Deactivated"
                        continue
            except:
                pass
    dropList = []
    for i,expiry in enumerate(expiryList):
        if expiry!='N/A':
            split = expiry.split('/')
            split[2] = split[2][:4]
            year = int(split[2])
            month = int(split[0])
            day = int(split[1])
            dayy = date(year,month,day)
            today = dayy.today()
            if (dayy-today).days>90:
                dropList.append(i)

    df["Status"] = statusList
    df["Expiry"] = mmcr.daysBetween(expiryList)

    df = df.drop(index=dropList)

    with pd.ExcelWriter("status2.xlsx",mode="w") as writer:
        df.to_excel(excel_writer = writer,index=False)