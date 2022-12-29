from selenium import webdriver
import re
import os
from selenium.webdriver.chrome.options import Options
import sys
from datetime import datetime as dt
from selenium.webdriver.common.keys import Keys
import datetime
import time
import xlrd
import pandas as pd

def download_xls():
    minus = dt.today().isoweekday()
    if minus < 7:
        startdate = dt.now() - datetime.timedelta(days=minus+7)
        smonth = startdate.month
        sday = startdate.day
        enddate = dt.now() - datetime.timedelta(days=minus)
        emonth = enddate.month
        eday = enddate.day
        startdate = str(startdate).split(' ', 1)
        print("The start date and time is", startdate[0])
        enddate = str(enddate).split(' ', 1)
        print("The end date and time is", enddate[0])
        print(f'{smonth}.{sday}~{emonth}.{eday}')
    else:
        startdate = dt.now() - datetime.timedelta(days=7)
        smonth = startdate.month
        sday = startdate.day
        enddate = dt.now()
        emonth = enddate.month
        eday = enddate.day
        startdate = str(startdate).split(' ', 1)
        print("The start date and time is", startdate[0])
        enddate = str(enddate).split(' ', 1)
        print("The end date and time is", enddate[0])
        print(f'{smonth}.{sday}~{emonth}.{eday}')

    chrome_options = Options()
    chrome_options.add_argument("--mute-audio")
    driver = webdriver.Chrome(options=chrome_options, executable_path="C:\\Users\\justin\\Desktop\\webdriver_selenium\\chromedriver.exe")

    driver.get('https://www.veding.com/static/v2/login.html')

    element = driver.find_element_by_xpath('//*[@id="id_username"]')
    element.send_keys('zaocan123')
    element = driver.find_element_by_xpath('//*[@id="id_password"]')
    element.send_keys('123456')
    driver.find_element_by_xpath('//*[@id="login_form"]/div[2]/div[4]/input').click()
    time.sleep(5)
    driver.find_element_by_xpath('//*[@id="left_scroller"]/div/div/dl[2]/dd/dl[3]').click()
    js = "document.getElementsByClassName('width-show')[1].style.display='block'"
    driver.execute_script(js)
    driver.find_element_by_xpath('//*[@id="shop_business_data9"]').click()

    driver.find_element_by_xpath('//*[@id="startDate"]').clear()
    element = driver.find_element_by_xpath('//*[@id="startDate"]')
    element.send_keys(startdate[0], ' 18:00:00')
    element.send_keys(Keys.ENTER)
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="endDate"]').clear()
    element = driver.find_element_by_xpath('//*[@id="endDate"]')
    element.send_keys(enddate[0], ' 18:00:00')
    element.send_keys(Keys.ENTER)
    time.sleep(1)
    # ----------------------------------------
    driver.find_element_by_xpath('//*[@id="selectType_chosen"]').click()
    element = driver.find_element_by_xpath('//*[@id="selectType_chosen"]/div/div/input')
    element.send_keys('店')
    element.send_keys(Keys.ENTER)
    time.sleep(3)
    # ----------------------------------------
    driver.find_element_by_xpath('//*[@id="shop_data_form"]/div[12]/button').click()
    time.sleep(3)
    driver.quit()
    return f'{smonth}.{sday+1}~{emonth}.{eday}'


def transfer_to_excel(file_path):
    sheet = pd.read_excel(io=file_path)
    new_sheet = pd.DataFrame(sheet)
    path = r"C:\\Users\\justin\\Desktop\\ask\\营收详情导出.xlsx"
    new_sheet.to_excel(path, sheet_name='Sheet1')

def find_path():
    pattern = re.compile(r'.+\.xls')
    at = []
    paththefile = []
    for root ,dirs,files in os.walk(r"C:\\Users\\justin\\Downloads\\"):
        for name in files:

            file_path=os.path.join(root,name)#包含路徑的檔案
            if pattern.search(file_path) is not None :
                #print(file_path)#匹配到的檔案 檔案路徑名

                ctime=time.localtime(os.path.getctime(file_path))#建立時間
                at.append(ctime)
                paththefile.append(file_path)
    print(paththefile[at.index(max(at))])
    return paththefile[at.index(max(at))]
