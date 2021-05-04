# -*- coding: utf-8 -*-
"""
Created on Tuesday 28-07-2020 15:30:09

@author: Nishanth T (Junior Python Developer)

"""

import os
import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.remote.webdriver import WebDriver as RemoteWebDriver
################################ Display Date ###########################
now = datetime.datetime.now()
date_time = now.strftime("%d-%b-%Y")
print(date_time)
filepath=r'D:/Nishu_works/Raw_data_web_scrapping/'+date_time
if not os.path.exists(filepath):
    os.makedirs(filepath)
Data=[]
################################ Reading excelsheet ###########################
with pd.ExcelFile(r"Excel sheet.xlsx") as reader:
    sheet = pd.read_excel(reader)
    for i in sheet:
        Data.append(sheet[i])
    S=Data[0].tolist()

################################ extract data ##########################
profile = webdriver.FirefoxProfile()
profile.set_preference("browser.download.dir",r'D:/Nishu_works/Raw_data_web_scrapping/');
profile.set_preference("browser.download.folderList",2);
profile.set_preference("browser.download.dir",r"D:/Nishu_works/Web_Scraping/"+date_time);
profile.set_preference("browser.helperApps.neverAsk.saveToDisk","application/zip")

url='Your website name'  # I have used a copyrighted Website so i cannot mention here. 

driver = webdriver.Firefox(firefox_profile=profile)
driver.get(url)

for j in range(0,len(S)):
   
    searchterm=S[j]
    sbox = driver.find_element_by_id("id")
    sbox.send_keys(searchterm)
    
    select = Select(driver.find_element_by_id("searchby"))
    select.select_by_visible_text("Visit ID")
    
    select = Select(driver.find_element_by_id("searchby"))
    select.select_by_visible_text("User ID")
    
    
    select = Select(driver.find_element_by_id("data"))
    select.select_by_visible_text("Original PDF")
    
    submit = driver.find_element_by_id("btn-download")
    submit.click()
    
    try:
        WebDriverWait(driver,4).until(EC.alert_is_present(),'waiting for popup to appear')  
        alert = driver.switch_to.alert
        alert.accept() 
        print("No Data Found for "+searchterm)                                                         
    except:
        print("Data is downloaded successfully "+searchterm)
        
    url=driver.command_executor._url
    print(url)
    session_id = driver.session_id 
    print(session_id)

    driver2 = webdriver.Remote(command_executor=url,desired_capabilities={})
    driver2.session_id = session_id
    searchterm=[]
    
