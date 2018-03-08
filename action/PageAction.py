#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@author: kyle
@time: 2018/3/6 16:23
"""
from selenium import webdriver
from util.ObjectMap import getElement
from util.ClipboardUtil import Clipboard
from util.KeyBoardUtil import KeyBoardKeys
from util.DirAndTime import *
from util.WaitUtil import WaitUtil
from selenium.webdriver.chrome.options import Options
import time


# 定义全局变量
driver = None
waitUtil = None


def open_browser(browserName, *arg):
    # 打开浏览器
    global driver, waitUtil
    try:
        if browserName.lower() == 'chrome':
            # 创建Chrome浏览器的一个Options实例对象
            chrome_options = Options()
            # 添加屏蔽 --ignore-certificate-errors提示信息的设置参数项
            chrome_options.add_experimental_option(
                "excludeSwitches",
                ["ignore-certificate-errors"]
            )
            driver = webdriver.Chrome(chrome_options=chrome_options)
        else:
            driver = webdriver.Firefox()
        waitUtil = WaitUtil(driver)
    except Exception as e:
        raise e


def visit_url(url, *arg):
    # 访问某个网址
    global driver
    try:
        driver.get(url)
    except Exception as e:
        raise e


def close_browser(*arg):
    # 关闭浏览器
    global driver
    try:
        driver.quit()
    except Exception as e:
        raise e


def sleep(sleepseconds, *arg):
    # 强制休眠
    try:
        time.sleep(sleepseconds)
    except Exception as e:
        raise e


def clear(locatorType, locatoeExpression, *arg):
    # 清除输入框中内容
    global driver
    try:
        getElement(driver, locatorType, locatoeExpression).clear()
    except Exception as e:
        raise e


def input_string(locatorType, locatorExpression, inputContent):
    # 输入框中输入数据
    global driver
    try:
        getElement(driver, locatorType, locatorExpression).send_keys(inputContent)
    except Exception as e:
        raise e


def click(locatorType, locatorExpression, *arg):
    # 单击页面元素
    global driver
    try:
        getElement(driver, locatorType, locatorExpression).click()
    except Exception as e:
        raise e


def assert_string_in_pagesource(assertString, *arg):
    # 断言元素在当前页面源码中
    global driver
    try:
        assert assertString in driver.page_source, u"%s not found in page source!" %assertString
    except AssertionError as e:
        raise AssertionError(e)
    except Exception as e:
        raise e


def assert_title(titleStr, *arg):
    # 断言标题中是否存在给定的关键字
    global driver
    try:
        assert titleStr in driver.title, u"%s not found in title!" %titleStr
    except AssertionError as e:
        raise AssertionError(e)
    except Exception as e:
        raise e


def getTitle(*arg):
    # 获取页面标题
    global driver
    try:
        driver.title
    except Exception as e:
        raise e


def getPageSource(*arg):
    # 获取页面源码
    global driver
    try:
        driver.page_source
    except Exception as e:
        raise e


def switch_to_frame(locatorType, locatorExpression, *arg):
    # 切换进frame
    global driver
    try:
        driver.switch_to.frame(getElement(driver, locatorType, locatorExpression))
    except Exception as e:
        raise e


def switch_to_default_content(*arg):
    # 切换出frame
    global driver
    try:
        driver.switch_to.default_content()
    except Exception as e:
        raise e


def paste_string(pasteString, *arg):
    # 模拟CRTL + V操作
    try:
        Clipboard.setText(pasteString)
        time.sleep(3)
        KeyBoardKeys.twoKeys('ctrl', 'v')
    except Exception as e:
        raise e


def press_tab_key(*arg):
    # 模拟Tab键
    try:
        KeyBoardKeys.oneKey('tab')
    except Exception as e:
        raise e


def press_enter_key(*arg):
    # 模拟enter键
    try:
        KeyBoardKeys.oneKey('enter')
    except Exception as e:
        raise e


def maximize_browser():
    # 窗口最大化
    global driver
    try:
        driver.maximize_window()
    except Exception as e:
        raise e


def capture_screen(*arg):
    # 截取屏幕图片
    global driver
    currTime = getCurrentTime()
    picNameAndPath = str(createCurrentDateDir()) + "/" + str(currTime) + ".png"
    try:
        driver.get_screenshot_as_file(picNameAndPath.replace('/', r'/'))
    except Exception as e:
        raise e
    else:
        return picNameAndPath


def waitPresenceOfElementLocated(locatorType, locatorExpression, *arg):
    """显示等待页面元素出现在DOM中，但不一定可见，存在则返回该页面元素对象"""
    global waitUtil
    try:
        waitUtil.presenceOfElementLocated(locatorType, locatorExpression)
    except Exception as e:
        raise e


def waitFrameToBeAvailableAndSwitchToIt(locatorType, locatorExpression, *arg):
    """检查frame是否存在，存在则切换进入frame控件中"""
    global waitUtil
    try:
        waitUtil.frameToBeAvailableAndSwitchToIt(locatorType, locatorExpression)
    except Exception as e:
        raise e


def waitVisibilityOfElementLocated(locatorType, locatorExpression, *arg):
    """显示等待页面元素出现在DOM中，并且可见，存在则返回该页面元素对象"""
    global waitUtil
    try:
        waitUtil.visibilityOfElementLocated(locatorType, locatorExpression)
    except Exception as e:
        raise e