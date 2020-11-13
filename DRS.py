import os

import pandas as pd
import numpy as np

from selenium import webdriver
from tkinter import messagebox
from selenium.webdriver.common.by import By

# setting up the path
os.chdir(os.path.dirname(os.path.realpath(__file__)))

# writing fuctions


def loadDataFromExcel():
    data2 = pd.read_excel("fullRecord.xlsx")
    #data2['Brand'] = data2['Brand'].str.lower()
    #data2['Brand'] = data2['Brand'].str.capitalize()
    return data2


def loginDRS():

    # login information
    email = 'Please put a valid username and password'
    password = 'please put a Password'

    # getting to the starting point
    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.get("https://dirbs.pta.gov.pk/drs")

    # waiting for the page to load
    driver.implicitly_wait(2)

    # logging in
    driver.find_element(By.ID, "identity").send_keys(email)
    # driver.find_element(By.ID, "login_password").send_keys(password)

    # wait to complete captcha
    messagebox.showinfo(title='Captcha', message='Finish CAPTCHA, and then click OK.')

    return driver


def getDeviceDetails(tempDriver, IMEI):
    tempDriver.get('https://dirbs.pta.gov.pk/drs/search/tac')

    # waiting for the page to load
    driver.implicitly_wait(1)

    tac = int(str(IMEI)[:8])

    tempDriver.find_element(By.ID, "search_tac").send_keys(tac)
    tempDriver.find_element(By.XPATH, "/html/body/div/div[1]/section[2]/div[2]/div[1]/form/div/span/button").click()

    # waiting for the page to load
    tempDriver.implicitly_wait(3)

    modelNumber = tempDriver.find_element(By.XPATH, "/html/body/div/div[1]/section[2]/div[2]/div[2]/table/tbody/tr[4]/td[2]").get_attribute('innerHTML')
    manufacturer = tempDriver.find_element(By.XPATH, '/html/body/div/div[1]/section[2]/div[2]/div[2]/table/tbody/tr[2]/td[2]').get_attribute('innerHTML')

    return modelNumber, manufacturer


def fillDeviceInfo(driver, df):
    noOfRows = df.shape[0]
    print('Number of rows', noOfRows)

    for row in range(noOfRows):
        print("print rows ", row)
        try:
            imeiNo = df.iloc[row, 3]
            deviceInfo = getDeviceDetails(driver, imeiNo)

            modelNumber = deviceInfo[0].replace('  ', '')
            manufacturer = deviceInfo[1].replace('  ', '')
            df.iloc[row, 6] = modelNumber
            df.iloc[row, 7] = manufacturer
        except:
            print('problem in row ', row)
            driver.implicitly_wait(3)
            driver.get("https://dirbs.pta.gov.pk/drs")

    print(df)


# Loading data into dataFrame
data = loadDataFromExcel()

# login DRS
driver = loginDRS()
driver.get('https://dirbs.pta.gov.pk/drs/search/tac')

# fill up status column with same value
data['Status'] = '1 IMEI register & 2nd is not'

# get device information from the IMEIs stored in panda
fillDeviceInfo(driver, data)

# resetting indexand making sure index starts at 1
data.reset_index(level=None, drop=True, inplace=True)
data.index = np.arange(1, len(data)+1)

# simplifying brand column
brandReplaceemnts = {'Samsung Korea': 'Samsung', 'Digicom Trading PVT Limited': 'Qmobile', 'HMD Global Oy': 'Nokia', 'Nokia Corporation': 'Nokia', 'Microsoft Mobile Oy': 'Nokia', 'Microsoft Mobile Oy, Nokia Corporation': 'Nokia'}
data['Brand'] = data['Brand'].replace(brandReplaceemnts)

data.to_excel('fullRecord 2.xlsx', index=True)
