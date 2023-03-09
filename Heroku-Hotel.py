# -*- coding: utf-8 -*-
"""
Created on Thu Mar  9 16:16:54 2023

@author: dayne
"""

from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.common.keys import Keys
import pandas as pd
import os
import time
import re
import os
from datetime import date
from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.common.keys import Keys
import pandas as pd
import os
import time
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By




today = date.today()
d10=str(today)
a10=re.findall(r"[0-9]{4}", d10)
a11=re.findall(r"[0-9]{2}", d10)
a12=a11[3]+'/'+a11[2]+'/'+a10[0]
os.chdir(r'C:\Users\dayne')
driver = webdriver.Chrome(ChromeDriverManager().install())

driver.get('https://app.axisrooms.com/supplier/arcHotelBookingReport.html')
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

Email = driver.find_element("id","emailId")
password = driver.find_element("id","password")

Email.send_keys("contact@summervillebeachresort.com")
password.send_keys("Mario12er@")

driver.find_element("xpath","//button[@class='g-recaptcha theme-button-one dark']").click()
time.sleep(4)
driver.get('https://app.axisrooms.com/supplier/arcHotelBookingReport.html')

driver.find_element("id","select2-searchType-container").click()
enter=driver.find_element("xpath","//input[@class='select2-search__field']")
enter.send_keys("Check In Date")
driver.find_element("xpath","//li[@class='select2-results__option select2-results__option--highlighted']").click()
start = driver.find_element("id","startdatehbr")
start.clear()
start.send_keys("01/04/2022")
end = driver.find_element("id","enddatehbr")
end.clear()
end.send_keys(a12)
driver.find_element("id","select2-selectstatus-container").click()
enter=driver.find_element("xpath","//input[@class='select2-search__field']")
enter.send_keys("Confirmed")
driver.find_element("xpath","//li[@class='select2-results__option select2-results__option--highlighted']").click()

driver.find_element("xpath","//a[@onclick='downloadForm();']").click()

import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials

scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
# add credentials to the account
creds = ServiceAccountCredentials.from_json_keyfile_name('beaming-inn-324610-b9aa26893b12.json', scope)

# authorize the clientsheet 
client = gspread.authorize(creds)
sheet = client.open('Hotel-Walk-In')

# get the first sheet of the Spreadsheet
sheet_instance = sheet.get_worksheet(0)

records_data = sheet_instance.get_all_records()

# view the data
records_data

records_df = pd.DataFrame.from_dict(records_data)
records_df.columns
# view the top records
records_df.head()
records_df1=records_df

import os
os.chdir(r'C:\Users\dayne\Downloads')

df = pd.read_excel('ArcBookingReport.xls', sheet_name='data')
df=df[['Product','Room Name (Count)','No of Rooms','Check In','Check Out','Net Amount','Channel']]

records_df=records_df[['Hotel Name','Room Type','Number Of Rooms','Check In','Check Out','Net Amount']]

df.rename(columns = {'Product':'Hotel Name', 'Room Name (Count)':'Room Type',
                              'No of Rooms':'Number Of Rooms'}, inplace = True)

nn10=[]
for i in range(0,len(records_df)):
    nn10.append('Walk-In')

records_df['Channel']=nn10    

df['fire1'] = df['Room Type'].str.split()
for i in range(0,len(df['fire1'])):
    df['Room Type'][i]=df['fire1'][i][0]+' '+df['fire1'][i][1]


del df['fire1']
nn1=[]
from datetime import datetime
for i in range(0,len(records_df)):
    nn1.append((datetime. strptime(records_df['Check Out'][i], '%Y-%m-%d')-datetime. strptime(records_df['Check In'][i], '%Y-%m-%d')).days)

nn2=[]
for i in range(0,len(df)):
    nn2.append((datetime. strptime(df['Check Out'][i], '%d/%m/%y')-datetime. strptime(df['Check In'][i], '%d/%m/%y')).days)

records_df['Number Of Nights']=nn1
df['Number Of Nights']=nn2

for i in range(0,len(df)):
    a13=re.findall(r"[0-9]{2}", df['Check In'][i])
    a14='20'+a13[2]+'-'+a13[1]+'-'+a13[0]
    df['Check In'][i]=a14
    
for i in range(0,len(df)):
    a13=re.findall(r"[0-9]{2}", df['Check Out'][i])
    a14='20'+a13[2]+'-'+a13[1]+'-'+a13[0]
    df['Check Out'][i]=a14


df_combined=pd.concat([df,records_df])

df_combined=df_combined.reset_index()
del df_combined['index']


df_combined['Average Rate Per Night']=df_combined['Net Amount']/df_combined['Number Of Nights']

os.chdir(r'C:\Users\dayne')

scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
# add credentials to the account
creds = ServiceAccountCredentials.from_json_keyfile_name('beaming-inn-324610-1d7ef5086ab6.json', scope)

# authorize the clientsheet 
client = gspread.authorize(creds)
sheet = client.open('Tableau')

from df2gspread import df2gspread as d2g
spreadsheet_key = '1fi9vyTDwWp0_BVuhen-3UdHpDQbXVBEpN-s0xzKZPbc'
wks_name = 'Sheet1'
d2g.upload(df_combined, spreadsheet_key, wks_name, credentials=creds, row_names=True)

os.chdir(r'C:\Users\dayne\Downloads')
os.remove("ArcBookingReport.xls")


os.chdir(r'C:\Users\dayne')
