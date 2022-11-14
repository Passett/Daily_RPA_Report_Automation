#This script was written by Richard Passett and downloads the files for Johnny from the Thompson Portal and sends them to Buck and Garrett. 
#It currently grabs 3 files using 3 different logins.

#Import dependencies
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import shutil
import keyring
from datetime import date, datetime, timedelta
import win32com.client
from zipfile import ZipFile
import re
import os

#Password variables for Grants Portal
GP_username=keyring.get_password("GP", "Reports username")
GP_password=keyring.get_password("GP", "Reports password")

#Directories used for script
holding_dir=r'J:\Admin & Plans Unit\Recovery Systems\3. Projects\Johnny_Automation\Holding_Folder'
attachment_destination=r'J:\Admin & Plans Unit\Recovery Systems\2. Reports\6. Other\Daily RPA Report\RPA Data'

#Grants Portal Listings used for script
RPA_Listing="https://grantee.fema.gov/#applicants/subrecipient?filters=1755840"

#Use webdriver for Chrome, set where you want the CSVs to download to, add other options/preferences as desired, point to where you have the driver downloaded, and set the driver to a variable.
#If you want to see what is happening in the browser, comment out the headless and disable-software-rasterizer options
options=webdriver.ChromeOptions()
prefs={
    'download.prompt_for_download': False,
    "download.default_directory" : holding_dir,
    'download.directory_upgrade': True,
    'plugins.always_open_pdf_externally': True
    }
options.add_experimental_option("prefs",prefs) 
options.add_experimental_option('excludeSwitches', ['enable-logging'])
#options.add_argument("--headless")
#options.add_argument("--disable-software-rasterizer")
options.add_argument("--start-maximized")
driver_service=Service(r"C:\Users\richardp\Desktop\chromedriver\chromedriver.exe")
driver=webdriver.Chrome(service=driver_service, options=options)
wait=WebDriverWait(driver, 120)

#Functions used for automation
#export report function
def download_GP_report():
    wait.until(EC.presence_of_element_located((By.CLASS_NAME,'caret')))
    wait.until(EC.element_to_be_clickable((By.CLASS_NAME,'caret')))
    dropdown_button=driver.find_element(By.CLASS_NAME,'caret')
    driver.execute_script("arguments[0].click();", dropdown_button)
    time.sleep(3)
    export_button=driver.find_element(By.XPATH,'//*[@id="accordion"]/div/div[1]/div[2]/div[2]/div/ul/li[5]/a')
    driver.execute_script("arguments[0].click();", export_button)

#function to move csv to desired destination. Waits for file to exist, empties destination folder before moving new file, and accounts for whether or not csv is in a zip file.
def move(destination):
    while len(os.listdir(holding_dir))==0: 
        time.sleep(10)
    for file in os.scandir(destination):
        os.remove(file.path)
    for item in os.listdir(holding_dir):
        file_name=holding_dir+"/"+item
        if item.endswith(".zip"):
            zip_ref = ZipFile(file_name) # create zipfile object
            zip_ref.extractall(destination) # extract file to dir
            zip_ref.close() # close file
            os.remove(file_name) #Delete original file
        elif item.endswith("crdownload"):
            time.sleep(10)
            move(destination)
        else:
            shutil.copy2(file_name, destination) #Copy csv to JDrive
            os.remove(file_name) #Delete original file
    time.sleep(5)

#Function to rename export file
def Rename_File(folder, file_name):
    for file in os.listdir(folder):
        old_file_name=folder+"/"+file
        if file.endswith(".csv"):
            new_file_name=folder+"/"+file_name+date.today().strftime("%m%d%Y")+".csv"
            os.rename(old_file_name, new_file_name)
        elif file.endswith(".xlsx"):
            new_file_name=folder+"/"+file_name+date.today().strftime("%m%d%Y")+".xlsx"
            os.rename(old_file_name, new_file_name)
        else:
            return

#Function to Export data from Grants Portal in CSV format. Accepts three arguments; listing (driver.get location) and destination (destination directory), and name (what you want the file to be named)
def GP_export(listing, destination, name):
    driver.get(listing)
    time.sleep(45)
    download_GP_report()
    move(destination)
    Rename_File(destination, name)

def email(to, cc):
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.Subject = 'Daily RPA Numbers'
    mail.HTMLBody = '<h3>Greetings,<br><br>Please see the attached daily RPA report.<br><br>Sincerely,<br><br>FDEM Recovery Admins</h3>'
    mail.Body = "Greetings,\r\n\r\nPlease see the attached daily RPA report.\r\n\r\nSincerely,\r\n\r\nFDEM Recovery Admins"
    mail.Attachments.Add(newfile)
    mail.CC = cc
    mail.Send()

print("Greetings, we will now export your data and email it to the appropriate contacts for you. We will provide updates along the way.")

#Open Grants Portal and login
driver.get("https://grantee.fema.gov/")
time.sleep(15)

#Part 1: Sign in
username_field=driver.find_element(By.ID,"username")
password_field=driver.find_element(By.ID,"password")
signIn_button=driver.find_element(By.ID,"credentialsLoginButton")
username_field.clear()
password_field.clear()
username_field.send_keys(GP_username)
password_field.send_keys(GP_password)
signIn_button.click()
time.sleep(15)
accept_button=driver.find_element(By.CSS_SELECTOR,"button.btn.btn-sm.btn-primary")
accept_button.click()
time.sleep(10)
accept_button2=driver.find_element(By.CSS_SELECTOR,"button.btn.btn-sm.btn-primary")
time.sleep(10)
accept_button2.click()
time.sleep(100)

#Part 2: Retreive passcode from email for authentication
outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")
inbox = mapi.GetDefaultFolder(6)
messages = inbox.Items
received_dt = datetime.now() - timedelta(minutes=5)
received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
messages = messages.Restrict("[SenderEmailAddress] = 'support.pagrants@fema.gov'")
messages = messages.Restrict("[Subject] = 'Grants Portal Request'")
for message in messages:
    text=message.Body

CodeRegexVariable=re.compile(r'(\d\d\d\d\d\d)')
code=CodeRegexVariable.search(str(text))
answer=code.group()

#Part 3: Enter code from email into Grants Portal and complete login
passcode_field=driver.find_element(By.ID,"passcode")
passcode_field.clear()
passcode_field.send_keys(answer)
submit_button=driver.find_element(By.ID,"otpSubmitButton")
submit_button.click()
time.sleep(60)
print("Successfully logged into Grants Portal")

#Download Report
GP_export(RPA_Listing, attachment_destination, "RPA Breakdown_")
print("RPA export successfully downloaded")
driver.close()

#Open the RPA report template, refresh the data sources and save as new name in correct location
print("prepping finalized report")
today=date.today().strftime("%m_%d_%Y")
filename=r'J:\Admin & Plans Unit\Recovery Systems\2. Reports\6. Other\Daily RPA Report\RPA_Breakdown.xlsx'
newfile=r'J:\Admin & Plans Unit\Recovery Systems\2. Reports\6. Other\Daily RPA Report\Final Report\\'+"RPA_Breakdown_"+today+".xlsx"

xl = win32com.client.DispatchEx("Excel.Application")
wb = xl.Workbooks.Open(filename)
xl.Visible = True
wb.RefreshAll()
xl.CalculateUntilAsyncQueriesDone()
time.sleep(15)
wb.SaveAs(newfile)
wb.Close(True)
xl.Quit()
print("Report completed")

#Open outlook and write email to Garrett and Buck, include subject, body, attachments
print("prepping emails")
email("kingofthe@road.com; trailerforsale@orrent.com", "roomstolendfor@50cents.com")
email("nophonenopool@nopets.com; Iaintgotno@cigarettes.com", "twohoursofpushingbroom@buysa8by124bedroom.com")
print("Emails sent, task complete!")
