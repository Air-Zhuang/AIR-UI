# _*_ coding: utf-8 _*_
# __author__ = 'Air Zhuang'

from testScripts.MainMethod import MainMethod
import os

if __name__ == '__main__':
    os.system('taskkill /F /IM ChromeDriver.exe')
    MainMethod()