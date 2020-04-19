# -*-coding:utf-8 -*-
# @Author : Zhigang

from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver
import time

def getElement(driver,locateType,locatorExpression):
    "获取单个页面元素对象"
    try:
        element=WebDriverWait(driver,5).until\
            (lambda x:x.find_element(by=locateType,value=locatorExpression))
        return element
    except Exception as e:
        raise e

def getElements(driver,locateType,locatorExpression):
    "获取多个页面元素对象，以list返回"
    try:
        elements=WebDriverWait(driver,5).until\
            (lambda x:x.find_elements(by=locateType,value=locatorExpression))
        return elements
    except Exception as e:
        raise e

if __name__=="__main__":
    driver=webdriver.Chrome(executable_path="D:\\chromedriver.exe")
    driver.get("http://mail.126.com")
    driver.implicitly_wait(10)
    driver.switch_to.frame(driver.find_element_by_xpath('//iframe[contains(@id,"x-URS-iframe")]'))
    searchBox=getElement(driver,"xpath",'//input[@name="email"]')
    searchBox.send_keys("123456")
    time.sleep(2)
    driver.quit()