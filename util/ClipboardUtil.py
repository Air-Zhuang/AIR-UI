# _*_ coding: utf-8 _*_
# __author__ = 'Air Zhuang'
# __date__ = '2018/6/27 22:10'

import win32clipboard as w
import win32con

class Clipboard(object):
    '''模拟Windows设置剪切板'''
    #读取剪贴板
    @staticmethod
    def getText():
        w.OpenClipboard()
        d=w.GetClipboardData(win32con.CF_TEXT)
        w.CloseClipboard()
        return d
    @staticmethod
    def setText(aString):
        w.OpenClipboard()
        w.EmptyClipboard()
        w.SetClipboardData(win32con.CF_UNICODETEXT,aString)
        w.CloseClipboard()
