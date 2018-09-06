# _*_ coding: utf-8 _*_
# __author__ = 'Air Zhuang'
# __date__ = '2018/6/26 21:37'

def cookieConvertFromStrToDict(str1):
    dict1={}
    list1=str1.split(";")
    for i in range(len(list1)):
        dict1[list1[i].split("=")[0]]=list1[i].split("=")[1]
    return dict1