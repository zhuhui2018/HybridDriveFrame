# -*-coding:utf-8 -*-
# @Author : Zhigang

import time
import os
from datetime import datetime
from config.VarConfig import *

def getCurrentDate():
    "获取当前的日期"
    #timeTup=time.localtime()
    #currentDate=str(timeTup.tm_year)+"-"+str(timeTup.tm_mon)+"-"+str(timeTup.tm_mday)
    currentDate=time.strftime("%Y-%m-%d",time.localtime())
    return currentDate

def getCurrentTime():
    "获取当前的时间"
    timeStr=datetime.now()
    # print (timeStr)
    nowTime=timeStr.strftime("%H-%M-%S-%f")
    # currentTime = time.strftime("%H-%M-%S", time.localtime())
    return nowTime

def createCurrentDateDir():
    "创建截图存放的目录"
    dirName=os.path.join(screenPicturesDir,getCurrentDate())
    if not os.path.exists(dirName):
        os.makedirs(dirName)
    return dirName

if __name__=="__main__":
    print (getCurrentDate())
    print (getCurrentTime())
    print (createCurrentDateDir())