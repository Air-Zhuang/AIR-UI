# _*_ coding: utf-8 _*_
# __author__ = 'Air Zhuang'
# __date__ = '2018/6/27 22:10'

from action.PageAction import *
from util.ParseExcel import ParseExcel
from config.VarConfig import *
import time
import traceback

#设置此次测试得环境编码为utf8
import sys
reload(sys)
sys.setdefaultencoding("utf-8")

excelObj=ParseExcel()
# excelObj.loadWorkBook(dataFilePath)

#用例或用例步骤执行结束后，向Excel中写执行结果信息
def writeTestResult(sheetObj,rowNo,colsNo,testResult,errorInfo=None,picPath=None):
    #测试通过结果信息为绿色，失败为红色
    colorDict={"pass":"green","faild":"red","interface":"purple"}
    #因为"测试用例"工作表和"用例步骤sheet表"中都有测试执行时间和测试结果列，定义此字典对象是为了区分具体应该写在哪个工作表
    colsDict={
        "testCase":[testCase_runTime,testCase_testResult],
        "caseStep":[testStep_runTime,testStep_testResult]
    }
    try:
        #在测试步骤sheet中，写入测试时间
        excelObj.writeCellCurrentTime(sheetObj,rowNo=rowNo,colsNo=colsDict[colsNo][0])
        #在测试步骤sheet中，写入测试结果
        if testResult in colorDict:
            excelObj.writeCell(sheetObj,content=testResult,rowNo=rowNo,colsNo=colsDict[colsNo][1],style=colorDict[testResult])
        else:
            #2.5版本：假如是接口测试，将状态码和返回结果写到测试结果列
            if testResult.startswith(("2"),):
                excelObj.writeCell(sheetObj,content=testResult,rowNo=rowNo,colsNo=colsDict[colsNo][1],style=colorDict["pass"])
            else:
                excelObj.writeCell(sheetObj, content=testResult, rowNo=rowNo, colsNo=colsDict[colsNo][1], style=colorDict["faild"])
        if errorInfo and picPath:
            #在测试步骤sheet中，写入异常信息
            excelObj.writeCell(sheetObj, content=errorInfo, rowNo=rowNo, colsNo=testStep_errorInfo)
            #在测试步骤sheet中，写入异常截图路径
            excelObj.writeCell(sheetObj, content=picPath, rowNo=rowNo, colsNo=testStep_errorPic)
        else:
            #在测试步骤sheet中，清空异常信息单元格
            excelObj.writeCell(sheetObj, content="", rowNo=rowNo, colsNo=testStep_errorInfo)
            # 在测试步骤sheet中，清空错误截图单元格
            excelObj.writeCell(sheetObj, content="", rowNo=rowNo, colsNo=testStep_errorPic)
    except Exception,e:
        print u"写excel出错",traceback.print_exc()
def MainMethod(excelNum=None):
    #多线程指定文件名
    global excelObj
    if excelNum==None:
        excelObj.loadWorkBook(dataFilePath)
    else:
        excelObj.loadWorkBook(dataFilePath2+str(excelNum))
    try:
        #根据Excel文件中的sheet名获取sheet对象
        caseSheet=excelObj.getSheetByName(u"测试用例")
        #获取测试用例sheet中是否执行列对象
        isExecuteColumn=excelObj.getColumn(caseSheet,testCase_isExecute)
        #记录执行成功的测试用例个数
        successfulCase=0
        #记录需要执行的用例个数
        requiredCase=0
        #运行之前清空case页的历史记录
        caseNum=excelObj.getRowsNumber(caseSheet)
        print u"-----------正在清空TestCase页的历史记录：--------------"
        # print "*"*(caseNum-2)
        for i in range(0,caseNum-1):
            # print "*",
            excelObj.writeCell2(caseSheet,content="",rowNo=i+2,colsNo=testCase_runTime)
            excelObj.writeCell2(caseSheet, content="", rowNo=i + 2, colsNo=testCase_testResult)
            excelObj.saveFile()
        print
        # 运行之前清空report页的历史记录
        print u"-----------正在清空Report页的历史记录：--------------"
        for i in range(2,caseNum+1):
            excelObj.writeCell2(excelObj.getSheetByName(u'自动化测试结果报表'),"",rowNo=report_title_row,colsNo=i)
            excelObj.writeCell2(excelObj.getSheetByName(u'自动化测试结果报表'), "", rowNo=report_pass_row, colsNo=i)
            excelObj.writeCell2(excelObj.getSheetByName(u'自动化测试结果报表'), "", rowNo=report_account_row, colsNo=i)
            excelObj.writeCell2(excelObj.getSheetByName(u'自动化测试结果报表'), "", rowNo=report_nopass_row, colsNo=i)
            excelObj.writeCell2(excelObj.getSheetByName(u'自动化测试结果报表'), "", rowNo=report_retio_row, colsNo=i)
            excelObj.saveFile()
            print
        for idx,i in enumerate(isExecuteColumn[1:]):
            #2.3版本bug,加了str()判断
            # if i.value.lower()=="y":
            if str(i.value).lower()=="y":
                requiredCase+=1
                caseRow=excelObj.getRow(caseSheet,idx+2)
                caseStepSheetName=caseRow[testCase_testStepSheetName-1].value
                # print caseStepSheetName
                stepSheet=excelObj.getSheetByName(caseStepSheetName)
                stepNum=excelObj.getRowsNumber(stepSheet)
                # print stepNum
                #在每个sheet执行之前清空它的历史记录
                if AutoDeleteLog.lower()=="yes":
                    print u"-----------正在清空[%s]的历史记录：--------------" % caseRow[testCase_testCaseName-1].value
                    print "*"*(stepNum-1)
                    for i in range(0,stepNum-1):
                        print "*",
                        excelObj.writeCell2(stepSheet,content="",rowNo=i+2,colsNo=testStep_runTime)
                        excelObj.writeCell2(stepSheet, content="", rowNo=i + 2, colsNo=testStep_testResult)
                        excelObj.writeCell2(stepSheet, content="", rowNo=i + 2, colsNo=testStep_errorInfo)
                        excelObj.writeCell2(stepSheet, content="", rowNo=i + 2, colsNo=testStep_errorPic)
                        excelObj.saveFile()
                successfulSteps=0
                print
                print u"-----------开始执行用例：%s--------------" % caseRow[testCase_testCaseName-1].value
                for step in xrange(2,stepNum+1):
                    stepRow=excelObj.getRow(stepSheet,step)
                    keyWord=stepRow[testStep_keyWords-1].value
                    locationType=stepRow[testStep_locationType-1].value
                    locationExpression=stepRow[testStep_locationExpression-1].value
                    operateValue=stepRow[testStep_operateValue-1].value
                    #坑！！！
                    if isinstance(operateValue,long):
                        operateValue=str(operateValue)
                    # print keyWord,locationType,locationExpression,operateValue
                    expressionStr=""
                    if keyWord and operateValue and locationType is None and locationExpression is None:
                        #坑！！！
                        if keyWord=="sleep":
                            expressionStr=keyWord.strip()+"("+operateValue+")"
                        elif keyWord=="changeHandle":
                            expressionStr=keyWord.strip()+"("+operateValue+")"
                        else:
                            expressionStr = keyWord.strip() + "(u'" + operateValue + "')"
                    elif keyWord and operateValue is None and locationType is None and locationExpression is None:
                        expressionStr=keyWord.strip()+"()"
                    elif keyWord and operateValue and locationType and locationExpression is None:
                        expressionStr = keyWord.strip() + "('"+locationType.strip()+"',u'"+operateValue+"')"
                    elif keyWord and operateValue and locationType and locationExpression:
                        #坑！！！
                        if keyWord=="posthascookie":
                            expressionStr = keyWord.strip() + "('" + locationType.strip()+"','"+locationExpression.replace("'",'"').strip() + "',u'" + str(operateValue)+"','"+cookie + "')"
                        else:
                            expressionStr = keyWord.strip() + "('" + locationType.strip() + "','" + locationExpression.replace("'",'"').strip() + "',u'" + str(operateValue)+ "')"
                    elif keyWord and locationType and locationExpression and operateValue is None:
                        if keyWord=="gethascookie":
                            expressionStr = keyWord.strip() + "('" + locationType.strip()+"','"+locationExpression.replace("'",'"').strip()+"','"+cookie + "')"
                        else:
                            expressionStr = keyWord.strip() + "('" + locationType.strip()+"','"+locationExpression.replace("'",'"').strip()+"')"
                    # print expressionStr
                    try:
                        eval(expressionStr)
                        excelObj.writeCellCurrentTime(stepSheet,rowNo=step,colsNo=testStep_runTime)
                    except Exception,e:
                        if keyWord=="posthascookie" or keyWord=="gethascookie" or keyWord=="postnocookie" or keyWord=="getnocookie":
                            errorInfo=traceback.format_exc()
                            writeTestResult(stepSheet,step,"caseStep","faild",errorInfo)
                            print u"步骤%s执行失败！" % stepRow[testStep_testStepDescribe-1].value
                            # print traceback.print_exc()
                        else:
                            capturePic=capture_screen()
                            errorInfo=traceback.format_exc()
                            writeTestResult(stepSheet, step, "caseStep", "faild", errorInfo,capturePic)
                            print u"步骤%s执行失败！" % stepRow[testStep_testStepDescribe-1].value
                            # print traceback.print_exc()
                            break
                    else:
                        #2.5版本：如果是post或get，将状态码写回Excel表格
                        if keyWord == "posthascookie" or keyWord == "gethascookie" or keyWord == "postnocookie" or keyWord == "getnocookie":
                            statusCode=str(eval(expressionStr).status_code)
                            returnText=str(eval(expressionStr).text)
                            writeTestResult(stepSheet, step, "caseStep",statusCode+"\n"+returnText)
                            successfulSteps+=1
                            print u"步骤%s执行通过！" % stepRow[testStep_testStepDescribe - 1].value
                        else:
                            writeTestResult(stepSheet, step, "caseStep", "pass")
                            successfulSteps += 1
                            print u"步骤%s执行通过！" % stepRow[testStep_testStepDescribe - 1].value
                '''报表'''
                excelObj.writeCell(excelObj.getSheetByName(u'自动化测试结果报表'), caseStepSheetName, rowNo=report_title_row, colsNo=idx+2)
                excelObj.writeCell(excelObj.getSheetByName(u'自动化测试结果报表'), successfulSteps, rowNo=report_pass_row, colsNo=idx + 2,style='green')
                excelObj.writeCell(excelObj.getSheetByName(u'自动化测试结果报表'), stepNum-1, rowNo=report_account_row, colsNo=idx + 2)
                excelObj.writeCell(excelObj.getSheetByName(u'自动化测试结果报表'), stepNum-1-successfulSteps, rowNo=report_nopass_row, colsNo=idx + 2,style='red')
                excelObj.writeCell(excelObj.getSheetByName(u'自动化测试结果报表'), float(format(float(successfulSteps)/float(stepNum-1),'.4f')), rowNo=report_retio_row,colsNo=idx + 2)
                #将结果写回case表
                if successfulSteps==stepNum-1:
                    writeTestResult(caseSheet,idx+2,"testCase","pass")
                    successfulCase+=1
                else:
                    writeTestResult(caseSheet, idx + 2, "testCase", "faild")
        print u"共 %d 条用例，%d 条需要被执行，本次执行通过 %d 条." % (len(isExecuteColumn)-1,requiredCase,successfulCase)
    except Exception,e:
        print traceback.print_exc()
if __name__ == '__main__':
    MainMethod()

















