# Mercedes

## How to Install
clone the repository using the following command once in the directory you choose you want it cloned to:
```git clone {insert http link}```
If you don't have git [here is a tutorial on how to download it.](https://git-scm.com/book/en/v2/Getting-Started-Installing-Git)
Once the repository is cloned we must ensure that the chromedriver.exe file in the dist folder is of the same version as your version of chrome.
For example, if you have chrome version 114 downloaded you must get chromedriver version 114. If you don't know how to check your version of chrome
[here is a tutorial](https://help.illinoisstate.edu/technology/support-topics/device-support/software/web-browsers/what-version-of-chrome-do-i-have).
Your best bet is to update chrome to the latest version and then download the latest version of the chromedriver. [Here is a link to the chromedriver downloads](https://chromedriver.chromium.org/downloads).
Once you download your new chromedriver unzip that file and find the file inside named chromedriver.exe, once you find it you can delete the chromedriver.exe file
that came with the git download and put the new one you downloaded in place of that (if you are on chrome v114 you can keep it in as it is a copy of chromedriver 114).
once you have you chromedriver in the dist file you should be able to 

## How The Scraper Works
The Eleads program collects the next two days service appointments and then filters out any car not between a 2019 and a 2023. It then goes and searces
the clients name in Eleads, if the scraper finds the associated vin from the service appointment schedule it grabs all email associated with the
account this information is then stored in the spreadsheet named "out.xlsx". The MMCR program searches MMCR for the emails found in the Eleads module in MMCR, upon 
finding an account successfully it will then click on the vehicle and view the status of the vehicles services, it then stores the status of the services and their 
expiration if necessary. Once all status' are obtained the resulting spreadsheet is saved to "satus.xlsx".
