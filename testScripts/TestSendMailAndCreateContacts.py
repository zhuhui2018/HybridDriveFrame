# -*-coding:utf-8 -*-
# @Author : Zhigang

from . import *
from testScripts.CreateContacts import *
from testScripts.WriteTestResult import *
from action.PageAction import  *
from util.Log import *
import traceback

def TestSendMailAndCreateContacts():
    try:
        # 根据excel文件中的sheet名获取sheet对象
        caseSheet=excelObj.getSheetByName("测试用例")
        # 获取测试用例sheet中是否执行列对象
        isExecuteColumn=excelObj.getColumn(caseSheet,testCase_isExecute)
        # 记录执行成功的测试用例个数
        successfulCase=0
        # 记录需要执行的用例个数
        requiredCase=0
        for id, data in enumerate(isExecuteColumn[1:]):
            # 因为用例sheet中第一行为标题行，无需执行，测试用例名称
            caseName=excelObj.getCellOfValue(caseSheet,rowNo=id+2,colsNo=testCase_testCaseName)
            # 循环遍历'测试用例'中的测试用例，执行被设置为执行的用例
            if data.value.lower() == "y":
                requiredCase += 1
                # 获取测试用例表中，用例执行时所使用的框架类型
                useFrameWorkName=excelObj.getCellOfValue(caseSheet,rowNo=id+2,colsNo=testCase_frameWorkName)
                # 获取测试用例表中，执行用例的步骤sheet名
                stepSheetName=excelObj.getCellOfValue(caseSheet,rowNo=id+2,colsNo=testCase_testStepSheetName)
                logging.info("--执行测试用例'%s'--" % caseName)
                if useFrameWorkName == "数据":
                    logging.info("******调用数据驱动******")
                    # 获取测试用例表中，执行框架为数据驱动的用例所使用的数据sheet名，colsNo为联系人
                    dataSheetName=excelObj.getCellOfValue(caseSheet,rowNo=id +2,colsNo=testCase_dataSourceSheetName)
                    # 获取测试用例的步骤sheet对象
                    stepSheetObj=excelObj.getSheetByName(stepSheetName)
                    # 获取测试用例使用的数据sheet对象
                    dataSheetObj=excelObj.getSheetByName(dataSheetName)
                    # 通过数据驱动框架执行添加联系人
                    result=dataDriverFun(dataSheetObj,stepSheetObj)
                    if result:
                        logging.info("用例'%s' 执行成功" % caseName)
                        successfulCase +=1
                        writeTestResult(caseSheet,rowNo=id+2,colsNo="testCase",testResult="pass")
                    else:
                        logging.info("用例'%s' 执行失败" % caseName)
                        writeTestResult(caseSheet,rowNo=id+2,colsNo="testCase",testResult="faild")
                elif useFrameWorkName == "关键字":
                    logging.info("******调用关键字驱动******")
                    caseStepObj=excelObj.getSheetByName(stepSheetName)
                    stepNums=excelObj.getRowsNumber(caseStepObj)
                    successfulSteps=0
                    logging.info("测试用例共'%s'步" % stepNums)
                    for index in range(2,stepNums+1):
                        # 第一行是标题行，无需执行
                        stepRow=excelObj.getRow(caseStepObj,index)
                        # 获取关键字作为调用的函数名
                        keyWord=stepRow[testStep_keywords-1].value
                        # 获取操作元素定位方式作为调用的函数的参数
                        locationType=stepRow[testStep_locationType-1].value
                        # 获取操作元素的定位表达式作为调用函数的参数
                        locatorExpression=stepRow[testStep_locatorExpression-1].value
                        operateValue = stepRow[testStep_operateValue - 1].value
                        if isinstance(operateValue, (int, float)):
                            operateValue = str(operateValue)
                        # PageAction.py文件中的页面动作函数的字符串表示
                        if (locationType and locatorExpression):
                            command = "%s('%s','%s','%s')" % (
                            keyWord, locationType, locatorExpression.replace("'", "\""),
                            operateValue) if operateValue else "%s('%s','%s')" % (
                                keyWord, locationType, locatorExpression.replace("'", "\""))
                        elif operateValue:
                            command = "%s(u'%s')" % (keyWord, operateValue)
                        else:
                            command = "%s()" % keyWord
                        #print(runStr)
                        try:
                            eval(command)
                        except Exception as e:
                            # 获取详细的异常堆栈信息
                            errorInfo=traceback.format_exc()
                            # 截取异常屏幕图片
                            capturePic=capture_screen()
                            writeTestResult(caseStepObj,rowNo=index,colsNo="testStep",testResult="faild"
                                                            ,errorInfo=str(errorInfo),picPath=capturePic)
                        else:
                            successfulSteps+=1
                            logging.info("执行步骤'%s'成功" % stepRow[testStep_testStepDescribe-1].value)
                            writeTestResult(caseStepObj,rowNo=index,colsNo="testStep",testResult="pass")
                    if successfulSteps == stepNums-1:
                        successfulCase += 1
                        logging.info("用例'%s'执行通过" % caseName)
                        writeTestResult(caseSheet,rowNo=id+2,colsNo="testCase",testResult="pass")
                    else:
                        logging.info("用例'%s'执行失败" % caseName)
                        writeTestResult(caseSheet,rowNo=id+2,colsNo="testCase",testResult="faild")
            else:
                # 清空不需要执行用例的执行时间和执行结果
                writeTestResult(caseSheet,rowNo=id+2,colsNo="testCase",testResult="")
                logging.info("用例'%s'被设置为忽略执行" % caseName)
        logging.info("共%s条用例，%s条需要被执行，成功执行%s条" % (len(isExecuteColumn)-1,requiredCase,successfulCase))
    except Exception as e:
        logging.debug("程序本身发生异常\n%s" % traceback.format_exc())


if __name__ == "__main__":
    TestSendMailAndCreateContacts()