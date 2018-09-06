# _*_ coding: utf-8 _*_
# __author__ = 'Air Zhuang'
# __date__ = '2018/6/27 22:08'

import os

'''是否开启自动删除测试记录功能'''
AutoDeleteLog="yes"

'''获取当前文件所在目录的父目录的绝对路径'''
parentDirPath=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

'''异常截图存放目录绝对路径'''
screenPicturesDir=parentDirPath+"\\exceptionpictures\\"

'''测试数据文件存放绝对路径'''
dataFilePath=parentDirPath+u"\\testData\\TestSuite.xlsx"
dataFilePath2=parentDirPath+u"\\testData\\"
processConfig=parentDirPath+u"\\testData\\ProcessConfig.txt"

'''测试数据文件中，测试用例表中部分列对应的数字序号'''
testCase_testCaseName=2
testCase_testStepSheetName=4
testCase_isExecute=5
testCase_runTime=6
testCase_testResult=7

'''测试步骤表中，部分列对应的数字序号'''
testStep_testStepDescribe=2
testStep_keyWords=3
testStep_locationType=4
testStep_locationExpression=5
testStep_operateValue=6
testStep_runTime=7
testStep_testResult=8
testStep_errorInfo=9
testStep_errorPic=10

'''报表表中，部分列对应的数字序号'''
report_title_row=1
report_pass_row=2
report_nopass_row=3
report_account_row=4
report_retio_row=5

'''cookie的字符串形式'''
cookie="_ha_id"