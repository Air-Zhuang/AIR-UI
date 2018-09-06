# _*_ coding: utf-8 _*_
# __author__ = 'Air Zhuang'
# __date__ = '2018/6/27 22:11'

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class WaitUtil(object):
    def __init__(self,driver):
        self.locationTypeDict={
            "xpath":By.XPATH,
            "id":By.ID,
            "name":By.NAME,
            "css_selector":By.CSS_SELECTOR,
            "class_name":By.CLASS_NAME,
            "tag_name":By.TAG_NAME,
            "link_text":By.LINK_TEXT,
            "partial_link_text":By.PARTIAL_LINK_TEXT
        }
        self.driver=driver
        self.wait=WebDriverWait(self.driver,30)
    def persenceOfElementLocated(self,locatorMethod,locatorExpression,*args):
        '''显式等待页面元素出现在DOM中，但并不一定可见，存在则返回该页面元素对象'''
        try:
            if self.locationTypeDict.has_key(locatorMethod.lower()):
                self.wait.until(
                    EC.presence_of_element_located((
                        self.locationTypeDict[locatorMethod.lower()],locatorExpression)))
            else:
                raise TypeError(u"未找到定位方式，请确认定位方法是否正确")
        except Exception,e:
            raise e
    def frameToBeAvailableAndSwitchToIt(self,locationType,locationExpression,*args):
        '''检查frame是否存在，存在则切换进frame控件中'''
        try:
            self.wait.until(
                EC.frame_to_be_available_and_switch_to_it((
                    self.locationTypeDict[locationType.lower()],locationExpression)))
        except Exception,e:
            raise e
    def visibilityOfElementLocated(self,locationType,locationExpression,*args):
        '''显示等待页面元素出现在DOM中，并且可见，存在则返回该页面元素对象'''
        try:
            self.wait.until(
                EC.visibility_of_element_located((
                    self.locationTypeDict[locationType.lower()],locationExpression)))
        except Exception,e:
            raise e
