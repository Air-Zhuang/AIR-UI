# _*_ coding: utf-8 _*_
# __author__ = 'Air Zhuang'
# __date__ = '2018/6/27 22:07'

from selenium import webdriver
from selenium.webdriver import ActionChains
from util.ObjectMap import getElement
from util.ClipboardUtil import Clipboard
from util.KeyBoardUtil import KeyboardKeys
from util.DirAndTime import *
from util.WaitUtil import WaitUtil
from util.cookieConvert import cookieConvertFromStrToDict
import time,os
import requests

#定义全局driver变量
driver=None
#全局的等待类实例对象
waitUtil=None

def postnocookie(headers,url,data):
    try:
        data2=(str(data))
        headers2=dict(eval(headers))
        res=requests.post(url=url,data=data2,headers=headers2,verify=False)
        return res
    except Exception,e:
        raise e
def getnocookie(headers,url):
    try:
        headers2=dict(eval(headers))
        res=requests.get(url=url,headers=headers2,verify=False)
        return res
    except Exception,e:
        raise e
def posthascookie(headers,url,data,cookie):
    try:
        data2=(str(data))
        headers2=dict(eval(headers))
        cookie2=cookieConvertFromStrToDict(cookie)
        res=requests.post(url=url,data=data2,headers=headers2,cookies=cookie2,verify=False)
        return res
    except Exception,e:
        raise e
def gethascookie(headers,url,cookie):
    try:
        headers2=dict(eval(headers))
        cookie2=cookieConvertFromStrToDict(cookie)
        res=requests.get(url=url,headers=headers2,cookies=cookie2,verify=False)
        return res
    except Exception,e:
        raise e
def open_browser(*args):
    global driver,waitUtil
    try:
        driver=webdriver.Chrome()
        waitUtil=WaitUtil(driver)
    except Exception,e:
        raise e
def open_browser_web(ip,*args):
    #使用AUTO-ISDP WEB端时将open_browser改为open_browser_web，并加上你的ip地址
    global driver,waitUtil
    try:
        driver=webdriver.Remote(
            command_executor='http://'+str(ip)+':8855/wd/hub',
            desired_capabilities={
                "browserName":"chrome",
                "video":"True",
                "platform":"WINDOWS"
            }
        )
    except Exception,e:
        raise e
def visit_url(url,*args):
    global driver
    try:
        driver.get(url)
    except Exception,e:
        raise e
def maximize_browser():
    global driver
    try:
        driver.maximize_window()
    except Exception, e:
        raise e
def implicitlyWait():
    global driver
    try:
        driver.implicitly_wait(25)
    except Exception, e:
        raise e
def refresh():
    global driver
    try:
        driver.refresh()
    except Exception, e:
        raise e
def close_browser(*args):
    global driver
    try:
        driver.quit()
    except Exception, e:
        raise e
def sleep(sleepSeconds,*args):
    try:
        time.sleep(sleepSeconds)
    except Exception, e:
        raise e
def clear(locationType,locationExpression,*args):
    global driver
    try:
        getElement(driver,locationType,locationExpression).clear()
        time.sleep(0.5)
    except Exception, e:
        raise e
def input_string(locationType,locationExpression,inputContent):
    global driver
    try:
        getElement(driver,locationType,locationExpression).send_keys(inputContent)
        time.sleep(0.5)
    except Exception, e:
        raise e
def script(value):
    global driver
    try:
        eval(value)
        time.sleep(0.5)
    except Exception,e:
        raise e
def click(locationType,locationExpression,*args):
    global driver
    try:
        getElement(driver,locationType,locationExpression).click()
        time.sleep(0.5)
    except Exception, e:
        raise e
def changeHandle(num,*args):
    global driver
    try:
        time.sleep(0.5)
        all_handles=driver.window_handles
        time.sleep(0.5)
        driver.switch_to_window(all_handles[num-1])
    except Exception,e:
        raise e
def assert_string_in_pagesource(assertString,*args):
    #断言页面源码是否存在某关键字或关键字符串
    global driver
    try:
        assert assertString in driver.page_source,u"%s not found in page source!" % assertString
    except AssertionError,e:
        raise AssertionError(e)
    except Exception,e:
        raise e
def assert_title(titleStr,*args):
    #断言页面标题是否存在给定的关键字符串
    global driver
    try:
        assert titleStr in driver.title,u"%s not found in title!" % titleStr
    except AssertionError,e:
        raise AssertionError(e)
    except Exception,e:
        raise e
def getTitle(*args):
    global driver
    try:
        return driver.title
    except Exception,e:
        raise e
def getPageSource(*args):
    global driver
    try:
        return driver.page_source
    except Exception,e:
        raise e
def switch_to_frame(locationType,frameLocatorExpression,*args):
    global driver
    try:
        driver.switch_to_frame(getElement(driver,locationType,frameLocatorExpression))
    except Exception,e:
        print "frame error"
        raise e
def switch_to_default_content(*args):
    global driver
    try:
        driver.switch_to_default_content()
    except Exception,e:
        raise e
def paste_string(pasteString,*args):
    #模拟ctrl+V操作
    try:
        Clipboard.setText(pasteString)
        time.sleep(1)
        KeyboardKeys.twoKeys("ctrl","v")
    except Exception,e:
        raise e
def press_tab_key(*args):
    try:
        time.sleep(0.5)
        KeyboardKeys.oneKey("tab")
    except Exception,e:
        raise e
def press_enter_key(*args):
    try:
        time.sleep(0.5)
        KeyboardKeys.oneKey("enter")
    except Exception,e:
        raise e
def press_del_key(*args):
    try:
        time.sleep(0.5)
        KeyboardKeys.oneKey("del")
    except Exception,e:
        raise e
def drag_and_drop_offset(locationType,locatorExpression,x_and_y):
    #拖拽页面元素指定坐标距离
    global driver
    try:
        l1=x_and_y.split('&')
        action_chains=ActionChains(driver)
        element=getElement(driver,locationType,locatorExpression)
        time.sleep(0.5)
        action_chains.drag_and_drop_by_offset(element,int(l1[0]),int(l1[1])).perform()
    except Exception,e:
        raise e
def doubleclick(locationType,locatorExpression):
    global driver
    try:
        action_chains=ActionChains(driver)
        element=getElement(driver,locationType,locatorExpression)
        time.sleep(0.5)
        action_chains.double_click(element).perform()
    except Exception,e:
        raise e
def rightclick(locationType,locatorExpression):
    global driver
    try:
        action_chains=ActionChains(driver)
        element=getElement(driver,locationType,locatorExpression)
        time,sleep(0.5)
        action_chains.context_click(element).perform()
    except Exception,e:
        raise e
def capture_screen(*args):
    global driver
    currTime=getCurrentTime()
    picNameAndPath=str(createCurrentDateDir())+"\\"+str(currTime)+".png"
    try:
        driver.get_screenshot_as_file(picNameAndPath.replace('\\',r'\\'))
    except Exception,e:
        raise e
    else:
        return picNameAndPath
















