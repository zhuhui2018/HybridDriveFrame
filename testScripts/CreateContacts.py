# -*-coding:utf-8 -*-
# @Author : Zhigang

from . import *
from testScripts.WriteTestResult import writeTestResult
from util.Log import *
from action.PageAction import *

def dataDriverFun(dataSourceSheetObj,stepSheetObj):
    """传入联系人数据sheet对象和创建联系人步骤sheet对象"""
    try:
        # 获取数据源表中是否执行列对象
        dataIsExecuteColumn=excelObj.getColumn(dataSourceSheetObj,dataSource_isExecute)
        # 获取数据源表中的"电子邮箱"对象，用于后续的断言
        emailColumn=excelObj.getColumn(dataSourceSheetObj,dataSource_email)
        # 获取测试步骤表中存在数据区域的行数
        stepRowNums=excelObj.getRowsNumber(stepSheetObj)
        # 记录成功执行的数据条数
        successDatas=0
        # 记录被设置为执行的数据条数
        requireDatas=0
        for id,data in enumerate(dataIsExecuteColumn[1:]):
            if data.value == "y":
                logging.info("开始添加联系人'%s'" % emailColumn[id+1].value)
                requireDatas+=1
                # 定义记录执行成功步骤数变量
                successStep=0
                for index in range(2,stepRowNums+1):
                    # 获取数据驱动测试步骤中第index行对象
                    rowObj=excelObj.getRow(stepSheetObj,index)
                    # 获取关键字作为调用的函数名 click
                    keyWord=rowObj[testStep_keywords-1].value
                    # 获取操作元素定位方式作为调用的函数的参数
                    locationType=rowObj[testStep_locationType-1].value
                    # 获取操作元素的定位表达式作为调用函数的参数
                    locatorExpression=rowObj[testStep_locatorExpression-1].value
                    # 获取操作值作为调用函数的参数
                    operateValue=rowObj[testStep_operateValue-1].value
                    if isinstance(operateValue,(int,float)):
                        operateValue=str(operateValue)
                    if operateValue and operateValue.isalpha():
                        # operateValue不为空，说明有操作值，从数据源表中根据坐标获取对应单元格的数据
                        coordinate=operateValue+str(id+2)
                        operateValue=excelObj.getCellOfValue(dataSourceSheetObj,coordinate=coordinate)
                    # 构造需要执行的python表达式，此表达式对应的是PageAction.py文件中的页面动作函数调用的字符串
                    if (locationType and locatorExpression):
                        command = "%s('%s','%s','%s')" % (
                            keyWord, locationType, locatorExpression.replace("'", "\""),
                            operateValue) if operateValue else "%s('%s','%s')" % (
                            keyWord, locationType, locatorExpression.replace("'", "\""))
                    elif operateValue:
                        command = "%s(u'%s')" % (keyWord, operateValue)
                    else:
                        command = "%s()" % keyWord
                    try:
                        # 通过eval函数，将拼接好的页面动作函数的字符串表示当成
                        # 有效的python表达式来执行，从而执行测试步骤sheet中关键字
                        # 在ageAction.py文件中对应的映射方法，来完成对页面元素的操作
                        if operateValue!="否":
                            # 当星标联系人为否时不会执行点击操作
                            eval(command)
                    except Exception as e:
                        logging.info("执行步骤'%s'发生异常" % rowObj[testStep_testStepDescribe-1].value)
                        raise e
                    else:
                        successStep+=1
                        logging.info("执行步骤'%s'成功" % rowObj[testStep_testStepDescribe-1].value)
                if stepRowNums == successStep +1:
                    successDatas += 1
                    # 如果成功执行的步骤数等于步骤表中给出的步骤数，说明第id+2
                    # 行的数据执行通过，写入通过信息
                    writeTestResult(sheetObj=dataSourceSheetObj,rowNo=id+2,colsNo="dataSheet",testResult="pass")
                else:
                    writeTestResult(sheetObj=dataSourceSheetObj,rowNo=id+2,colsNo="dataSheet",testResult="faild")
            else:
                # 将不需要执行的数据行的执行时间和执行结果单元格清空
                writeTestResult(sheetObj=dataSourceSheetObj,rowNo=id+2,colsNo="dataSheet",testResult="")
        if requireDatas == successDatas:
            # 当成功执行的数据条数等于被设置为需要执行的数据条数，才表示调用数据驱动的测试用例执行通过
            return 1
        else:
            # 表示调用数据驱动的测试用例执行失败
            return 0
    except Exception as e:
        raise e

