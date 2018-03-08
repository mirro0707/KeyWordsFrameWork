#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@author: kyle
@time: 2018/3/6 15:57
"""
import os,time
from datetime import datetime
from config.VarConfig import screenPicturesDir


# 获取当前日期
def getCurrentDate():
    date = time.localtime()
    currentDate = str(date.tm_year) + "-" + str(date.tm_mon) + "-" + str(date.tm_mday)
    return currentDate


# 获取当前时间
def getCurrentTime():
    now = datetime.now()
    nowTime = now.strftime('%H-%M-%S-%f')
    return nowTime

# 创建截图保存目录
def createCurrentDateDir():
    dirName = os.path.join(screenPicturesDir, getCurrentDate())
    if not os.path.exists(dirName):
        os.makedirs(dirName)
    return dirName
