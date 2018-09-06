# _*_ coding: utf-8 _*_
# __author__ = 'Air Zhuang'

from testScripts.MainMethod import MainMethod
import os
from multiprocessing import Process
from config.VarConfig import processConfig

if __name__ == '__main__':
    os.system('taskkill /F /IM ChromeDriver.exe')
    threads=[]
    with open(processConfig,'rb+') as f:
        lines=f.readlines()
        print u"正在同时执行："
        for i in lines:
            print i
        for i in lines:
            t=Process(target=MainMethod,args=(str(i).strip(),))
            threads.append(t)
        for i in threads:
            i.start()
        for i in threads:
            i.join()