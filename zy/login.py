#!/usr/bin/python
# -*- coding: UTF-8 -*-

#@Authod : wuyanhui

'''
登陆
'''
from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains


URL = 'http://132.121.92.224:8813/telant'
ACCOUNTID = 'userName'
PASSWORDID = 'userPass'
ACCOUNT = 'dushuiyun'
PASSWORD = 'Abc123456**'
BTNSUBMIT = 'btnSubmit'

URL_GX = 'http://132.121.92.224:8813/BundleSmart/Main.html?userToken=4386995106'

def main():
    print('main')
    # driver = webdriver.Chrome(executable_path='chromedriver.exe')
    # driver = webdriver.Edge(executable_path='MicrosoftWebDriver14393.exe')
    # driver = webdriver.Ie(executable_path='IEDriverServer.exe')
    driver = webdriver.Firefox(executable_path='geckodriver.exe')
    driver.get(URL)
    # driver.get(' https://get.adobe.com/flashplayer/')
    driver.implicitly_wait(10)
    driver.find_element_by_id(ACCOUNTID).send_keys(ACCOUNT)
    driver.find_element_by_id(PASSWORDID).send_keys(PASSWORD)
    driver.find_element_by_id(BTNSUBMIT).click()
    driver.implicitly_wait(10)

    driver.get(URL_GX)
    # driver.implicitly_wait(100)
    time.sleep(5)
    # driver.implicitly_wait(10)
    ActionChains(driver).move_by_offset(xoffset=236,yoffset=55).context_click().perform()


    time.sleep(600)
    driver.close()
    driver.quit()

if __name__=='__main__':
    main()

