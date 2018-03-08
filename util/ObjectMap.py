#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@author: kyle
@time: 2018/2/28 15:01
"""
from selenium.webdriver.support.ui import WebDriverWait


# 获取单个页面元素对象
def getElement(driver, locateType, locateExpression):
    try:
        element = WebDriverWait(driver, 30).until(lambda x :x.find_element(by=locateType, value=locateExpression))
        return element
    except Exception as e:
        raise e

# 获取多个相同页面元素对象，list形式返回
def getElements(driver, locateType, locateExpression):
    try:
        elements = WebDriverWait(driver, 30).until(lambda x :x.find_elements(by=locateType, value=locateExpression))
        return elements
    except Exception as e:
        raise e