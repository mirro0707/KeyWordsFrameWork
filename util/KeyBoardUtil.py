#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@author: kyle
@time: 2018/3/6 14:50
"""
import win32con
import win32api


class KeyBoardKeys(object):
    '''
    模拟键盘按键类
    '''
    VK_CODE = {
        'enter': 0x0D,
        'ctrl': 0x11,
        'v' : 0x56
    }

    @staticmethod
    def keyDown(keyName):
        # 按下按键
        win32api.keybd_event(KeyBoardKeys.VK_CODE[keyName], 0, 0, 0)

    @staticmethod
    def keyUp(keyName):
        # 释放按键
        win32api.keybd_event(KeyBoardKeys.VK_CODE[keyName], 0, win32con.KEYEVENTF_KEYUP, 0)

    @staticmethod
    def oneKey(key):
        # 模拟单个按键
        KeyBoardKeys.keyDown(key)
        KeyBoardKeys.keyUp(key)

    @staticmethod
    def twoKeys(key1, key2):
        # 模拟两个组合按键
        KeyBoardKeys.keyDown(key1)
        KeyBoardKeys.keyDown(key2)
        KeyBoardKeys.keyUp(key2)
        KeyBoardKeys.keyUp(key1)