# -*-coding:utf-8 -*-
# @Author : Zhigang

import logging.config
from config.VarConfig import parentDirPath

#读取日志的配置文件
# print (parentDirPath)
logging.config.fileConfig(parentDirPath + "\\config\\Logger.conf")

#选择一个日志格式
logger=logging.getLogger("example02")

def debug(message):
    "打印debug级别的信息"
    logger.debug(message)

def info(message):
    "打印info级别的信息"
    logger.info(message)

def warning(message):
    "打印warnning级别的信息"
    logger.warning(message)

if __name__=="__main__":
    print ("conf file path:",parentDirPath+"\\config\\Logger.conf")
    info("hi")
    debug("world")
    warning("gloryroad")