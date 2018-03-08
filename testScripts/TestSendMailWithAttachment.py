#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@author: kyle
@time: 2018/3/6 15:03
"""
from action.PageAction import *
from util.ParseExcel import ParseExcel
from config.VarConfig import *
import traceback


# 创建解析excel对象
excelObj =  ParseExcel()
# 将excel数据文件加载到内存中
excelObj.loadWorkBook(dataFilePath)


# 用例执行结束后，向excel中写执行结果信息
def writeTestResult(sheetObj, rowNo, colNo, testResult, errorInfo=None, picPath=None):
    # 定义颜色字典
    colorDict = {"pass":"green", "fail":"red"}
    # 区分工作表
    colsDict = {
        "testCase" : [testCase_runTime, testCase_testResult],
        "caseStep" : [testStep_runTime, testStep_testResult]
    }
    try:
        # 测试步骤sheet中，写入测试时间
        excelObj.writeCellCurrentTime(sheetObj, rowNo=rowNo, colNo=colsDict[colNo][0])
        # 测试步骤sheet中，写入测试结果
        excelObj.writeCell(sheetObj, content=testResult, rowNo=rowNo, colNo=colsDict[colNo][1],
                           style=colorDict[testResult])
        if errorInfo and picPath:
            # 在测试步骤sheet页中写入异常信息
            excelObj.writeCell(sheetObj, content=errorInfo, rowNo=rowNo, colNo=testStep_errorInfo)
            # 在测试步骤sheet中，写入异常截图路径
            excelObj.writeCell(sheetObj, content=picPath, rowNo=rowNo, colNo=testStep_errorPic)
        else:
            # 测试步骤sheet中，清空异常信息单元格
            excelObj.writeCell(sheetObj, content="", rowNo=rowNo, colNo=testStep_errorInfo)
            # 测试步骤sheet中，清空异常截图单元格
            excelObj.writeCell(sheetObj, content="", rowNo=rowNo, colNo=testStep_errorPic)
    except Exception as e:
        print(u"写excel出错，",traceback.print_exc())


def TestSendMailWithAttachment():
    try:
        # 根据excel文件中的sheet名获取sheet对象
        caseSheet = excelObj.getSheetByName(u"testcase")
        # 获取测试用例sheet中是否执行（isExecute）对象
        isExecuteColumn = excelObj.getColumn(caseSheet, testCase_isExecute)
        # 记录执行成功的测试用例个数
        successfulCase = 0
        # 记录执行失败的测试用例个数
        requiredCase = 0
        for idx, i in enumerate(isExecuteColumn[1:]):
            # 循环遍历“testcase”表中的用例，执行设置为执行的测试用例
            if i.value.lower() == "y":
                requiredCase += 1
                # 获取“testcase”表中第idx + 2行数据
                caseRow = excelObj.getRow(caseSheet, idx+2)
                # 获取第idx+2行数据的单元格内容
                caseStepSheetName = caseRow[testCase_testStepSheetName - 1].value
                # 根据用例步骤获取步骤sheet对象
                stepSheet = excelObj.getSheetByName(caseStepSheetName)
                # 获取步骤sheet中步骤数
                stepNum = excelObj.getRowNumber(stepSheet)
                # 记录测试用例i的步骤成功数
                successfulSteps = 0
                print(u"开始执行用例%s") %caseRow[testCase_testCaseName - 1].value
                for step in range(2, stepNum+1):
                    # 获取步骤sheet中第step行对象
                    stepRow = excelObj.getRow(stepSheet, step)
                    # 获取关键字作为调用的函数名
                    keyWord = stepRow[testStep_keyWords - 1].value
                    # 获取操作元素定位方式作为调用的函数的参数
                    locatorType = stepRow[testStep_locatorType - 1].value
                    # 获取操作元素定位表达式作为调用的函数的参数
                    locatorExpression = stepRow[testStep_locatorExpression - 1].value
                    # 获取操作值
                    operateValue = stepRow[testStep_operateValue - 1].value

                    # 将int型数据转换成string型，便于字符串拼接
                    if isinstance(operateValue, int):
                        operateValue = str(operateValue)

                    expressionStr = ""
                    # 构造需要执行的python语句
                    # 对应的是PageAction.py文件中的页面动作函数调用的字符串表示
                    if keyWord and operateValue and locatorType is None and locatorExpression is None:
                        expressionStr = keyWord.strip() + "(u'" + operateValue + "')"
                    elif keyWord and operateValue is None and locatorType is None and locatorExpression is None:
                        expressionStr = keyWord.strip() + "()"
                    elif keyWord and locatorType and operateValue and locatorExpression is None:
                        expressionStr = keyWord.strip() + "('" + locatorType.strip() + "',u'" + operateValue + "')"
                    elif keyWord and locatorType and locatorExpression and operateValue:
                        expressionStr = keyWord.strip() + "('" + locatorType.strip() + \
                                        "', '" + locatorExpression.replace("'", '"').strip() + \
                                        "',u'" + operateValue + "')"
                    elif keyWord and locatorType and locatorExpression and operateValue is None:
                        expressionStr = keyWord.strip() + "('" + locatorType.strip() + "', '" + \
                                        locatorExpression.replace(",", '"').strip() + "')"
                    try:
                        # 通过eval函数，将拼接的页面动作函数调用的字符串表示
                        # 当成有效的python表达式执行，从而执行测试步骤sheet中关键字在Action.py文件中对应的映射方法，
                        # 来完成页面操作
                        eval(expressionStr)
                        # 在测试执行时间列写入执行时间
                        excelObj.writeCellCurrentTime(stepSheet, rowNo=step, colNo=testStep_runTime)
                    except Exception as e:
                        # 截取异常图片
                        capturePic = capture_screen()
                        # 获取详细堆栈信息
                        errorInfo = traceback.format_exc()
                        # 在测试步骤sheet页中写入失败信息
                        writeTestResult(stepSheet, step, "caseStep", "faild", errorInfo, capturePic)
                        print(u"步骤“{0}”执行失败".format(stepRow[testStep_testStepDescribe - 1].value))
                    else:
                        # 在测试步骤sheet中写入成功信息
                        writeTestResult(stepSheet, step, "caseStep", "pass")
                        # 每成功一步，successfulStep加1
                        successfulSteps += 1
                        print(u"步骤“{0}”执行通过".format(stepRow[testStep_testStepDescribe - 1].value))
                if successfulSteps == stepNum -1:
                    # 当测试用例步骤sheet中所有步骤都执行成功，认为此测试用例执行通过，将成功信息写入测试用例工作表中
                    # 否则写入失败信息
                    writeTestResult(caseSheet, idx+2, "testCase", "pass")
                    successfulCase += 1
                else:
                    writeTestResult(caseSheet, idx+2, "testCase", "faild")
        print(u"共有{0}条用例，{1}条需要被执行，本次执行通过{2}条".format
              (len(isExecuteColumn) -1, requiredCase, successfulCase))
    except Exception as e:
        # 打印详细异常堆栈信息
        print(traceback.print_exc())
