# -*-coding:utf-8 -*-
# @Author : Zhigang

from util.ParseExcel import ParseExcel
from config.VarConfig import *


"创建解析Excel对象"
excelObj=ParseExcel()
"将excel数据文件加载到内存"
excelObj.loadWorkBook(dataFilePath)
