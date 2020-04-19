# HybridDriveFrame
基于混合驱动框架完成对126邮件的登录、添加联系人信息、自动发送邮件操作


action
    
    --PageAction.py    关键字函数

config

    --Logger.conf   日志配置文件
    --VarConfig.py 工程路径

exceptionpictures

    存放异常截图

log 

    存放日志文件

testData

    只需更改excel文件中的登录sheet表用户名和密码；增加、删除、修改联系人即可

testScripts

    --CreateContacts.py   创建联系人模块
    --TestSendMailAndCreateContacts.py      解析testData，完成发送邮件并创建联系人操作，主脚本
    --WriteTestResult.py    记录测试结果

util

    --ClipboardUtil.py     模拟设置剪切板
    --DirAndTime.py        时间模块
    --KeyBoardUtil.py      模拟键盘按键类
    --Log.py                      日志模块
    --ObjectMap.py         获取页面元素
    --ParseExcel.py           读写excel模块
    --WaitUtil.py               显示等待模块

RunTest.py    主程序入口，操作此脚本可完成程序运行
 
未封装的测试代码，可忽略：

    add_contact_send_mail.py
    AddContactSendEmailWithAttachment.py
    TestSendMailWithAttachment.py
