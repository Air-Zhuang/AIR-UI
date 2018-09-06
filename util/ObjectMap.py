# _*_ coding: utf-8 _*_
# __author__ = 'Air Zhuang'
# __date__ = '2018/6/27 22:11'

from selenium.webdriver.support.ui import WebDriverWait

def getElement(driver,locateType,locatorExpression):
    element=None
    if str(locateType).lower()=='id':
        element=driver.find_element_by_id(str(locatorExpression))
    if str(locateType).lower()=='xpath':
        element=driver.find_element_by_xpath(str(locatorExpression))
    if str(locateType).lower()=='classname':
        element=driver.find_element_by_class_name(str(locatorExpression))
    if str(locateType).lower()=='name':
        element=driver.find_element_by_name(str(locatorExpression))
    if str(locateType).lower()=='linktext':
        element=driver.find_element_by_link_text(str(locatorExpression))
    if str(locateType).lower()=='partiallinktext':
        element=driver.find_element_by_partial_link_text(str(locatorExpression))
    if str(locateType).lower()=='tagname':
        element=driver.find_element_by_tag_name(str(locatorExpression))
    if str(locateType).lower()=='css':
        element=driver.find_element_by_css_selector(str(locatorExpression))
    if locateType.endswith(('0','1','2','3','4','5','6','7','8','9')):
        locateType2=locateType[:-1]
        num=locateType[-1]
        if str(locateType2).lower() == 'id':
            element = driver.find_elements_by_id(str(locatorExpression))[int(num)-1]
        if str(locateType2).lower() == 'xpath':
            element = driver.find_elements_by_xpath(str(locatorExpression))[int(num)-1]
        if str(locateType2).lower() == 'classname':
            element = driver.find_elements_by_class_name(str(locatorExpression))[int(num)-1]
        if str(locateType2).lower() == 'name':
            element = driver.find_elements_by_name(str(locatorExpression))[int(num)-1]
        if str(locateType2).lower() == 'linktext':
            element = driver.find_elements_by_link_text(str(locatorExpression))[int(num)-1]
        if str(locateType2).lower() == 'partiallinktext':
            element = driver.find_elements_by_partial_link_text(str(locatorExpression))[int(num)-1]
        if str(locateType2).lower() == 'tagname':
            element = driver.find_elements_by_tag_name(str(locatorExpression))[int(num)-1]
        if str(locateType2).lower() == 'css':
            element = driver.find_elements_by_css_selector(str(locatorExpression))[int(num)-1]
        if locateType2.endswith(('0', '1', '2', '3', '4', '5', '6', '7', '8', '9')):
            locateType3 = locateType[:-2]
            num = locateType[-2:]
            if str(locateType3).lower() == 'id':
                element = driver.find_elements_by_id(str(locatorExpression))[int(num) - 1]
            if str(locateType3).lower() == 'xpath':
                element = driver.find_elements_by_xpath(str(locatorExpression))[int(num) - 1]
            if str(locateType3).lower() == 'classname':
                element = driver.find_elements_by_class_name(str(locatorExpression))[int(num) - 1]
            if str(locateType3).lower() == 'name':
                element = driver.find_elements_by_name(str(locatorExpression))[int(num) - 1]
            if str(locateType3).lower() == 'linktext':
                element = driver.find_elements_by_link_text(str(locatorExpression))[int(num) - 1]
            if str(locateType3).lower() == 'partiallinktext':
                element = driver.find_elements_by_partial_link_text(str(locatorExpression))[int(num) - 1]
            if str(locateType3).lower() == 'tagname':
                element = driver.find_elements_by_tag_name(str(locatorExpression))[int(num) - 1]
            if str(locateType3).lower() == 'css':
                element = driver.find_elements_by_css_selector(str(locatorExpression))[int(num) - 1]
    return element