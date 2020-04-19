# -*-coding:utf-8 -*-
# @Author : Zhigang

import win32clipboard as w
import win32con

class Clipboard(object):
    """模拟windows设置剪切板"""

    "读取剪切板"
    @staticmethod
    def getText():
        # 打开剪切板
        w.OpenClipboard()
        # 获取剪切板中的数据
        d=w.GetClipboardData(win32con.CF_TEXT)
        # 关闭剪切板
        w.CloseClipboard()
        # 返回剪切板数据给调用者
        return d

    "设置剪切板内容"
    @staticmethod
    def setText(aString):
        # 打开剪切板
        w.OpenClipboard()
        # 清空剪切板
        w.EmptyClipboard()
        # 将数据aString写入剪切板
        w.SetClipboardData(win32con.CF_UNICODETEXT,aString)
        # 关闭剪切板
        w.CloseClipboard()


if __name__=="__main__":
    Clipboard.setText("123")
    print (type(Clipboard.getText()))
    p=Clipboard()
    p.setText("上班")
    print (p.getText().decode("gbk"))