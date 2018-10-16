#!/usr/bin/env python
#-*- coding:utf-8 -*-
# coding=utf-8
import unittest
import time
import HTMLTestRunner
from email.mime.text import MIMEText   #邮件文本类
from email.mime.multipart import MIMEMultipart  #邮件附件需要的类
import smtplib    #用来发送邮件
import os

####下面三行代码python2报告出现乱码时候可以加上####
import sys
reload(sys)
sys.setdefaultencoding('utf8')

# 这个是优化版执行所有用例并发送报告，分四个步骤
# 第一步加载用例
# 第二步执行用例
# 第三步获取最新测试报告
# 第四步发送邮箱 （这一步不想执行的话，可以注释掉最后面那个函数就行）

# 当前脚本所在文件真实路径
cur_path = os.path.dirname(os.path.realpath(__file__))

def add_case(caseName="case", rule="test*.py"):
    '''第一步：加载所有的测试用例'''
    case_path = os.path.join(cur_path, caseName)  # 用例文件夹
    # 如果不存在这个case文件夹，就自动创建一个
    if not os.path.exists(case_path):os.mkdir(case_path)
    print("test case path:%s"%case_path)
    # 定义discover方法的参数
    discover = unittest.defaultTestLoader.discover(case_path,
                                                  pattern=rule,
                                                  top_level_dir=None)
    """
    discover 方法里面有三个参数:
    -case_dir:这个是待执行用例的目录。
    -pattern:这个是匹配脚本名称的规则，test*.py 意思是匹配 test 开头的所有脚本。 
    -top_level_dir:这个是顶层目录的名称，一般默认等于 None 就行了。
    """
    print(discover)
    return discover


def run_case(all_case, reportName="report"):
    '''第二步：执行所有的用例, 并把结果写入HTML测试报告'''
    now = time.strftime("%Y_%m_%d_%H_%M_%S")
    report_path = os.path.join(cur_path, reportName)  # 用例文件夹
    # 如果不存在这个report文件夹，就自动创建一个
    if not os.path.exists(report_path):os.mkdir(report_path)
    report_abspath = os.path.join(report_path, now+"result.html")
    print("report path:%s"%report_abspath)
    fp = open(report_abspath, "wb")
    runner = HTMLTestRunner.HTMLTestRunner(stream=fp,
                                           title=u'自动化测试报告，测试结果如下：',
                                           description=u'用例执行情况：')
    # 调用add_case函数返回值
    runner.run(all_case)
    fp.close()

def get_report_file(report_path):
    '''第三步：获取最新的测试报告'''
    """
    sort按key的关键字进行升序排序，lambda的入参fn为lists列表的元素，获取文件的最后修改时间，所以最终以文件时间从小到大排序
    os.path.join 将多个路径拼接
    """
    lists = os.listdir(report_path)    #os.listdir() 方法用于返回指定的文件夹包含的文件或文件夹的名字的列表。这个列表以字母顺序
    lists.sort(key=lambda fn: os.path.getmtime(os.path.join(report_path, fn)))
    print (u'最新测试生成的报告： '+lists[-1])
    # 找到最新生成的报告文件
    report_file = os.path.join(report_path, lists[-1])
    return report_file

def send_mail(sender, psw, receiver, smtpserver, report_file, port):
    '''第四步：发送最新的测试报告内容'''
    with open(report_file, "rb") as f:
        mail_body = f.read()                 # read() 函数，该函数会会一次性读取文件的全部内容
    # 定义邮件内容
    # 创建一个带附件的实例
    msg = MIMEMultipart()
    body = MIMEText(mail_body, _subtype='html', _charset='utf-8')  # 三个参数：第一个为文本内容，第二个 设置文本格式，第三个 utf-8 设置编码
    msg['Subject'] = u"自动化测试报告"
    msg["from"] = sender
    msg["to"] = psw
    msg.attach(body)  #邮件正文部分
    # 添加附件
    att = MIMEText(open(report_file, "rb").read(), "base64", "utf-8")
    att["Content-Type"] = "application/octet-stream"
    att["Content-Disposition"] = 'attachment; filename= "report.html"'
    msg.attach(att)
    try:
        smtp = smtplib.SMTP_SSL(smtpserver, port)
    except:
        smtp = smtplib.SMTP()
        smtp.connect(smtpserver, port)
    # 用户名密码
    smtp.login(sender, psw)
    smtp.sendmail(sender, receiver, msg.as_string())
    smtp.quit()
    print('test report email has send out !')


if __name__ == "__main__":
    all_case = add_case()   # 1加载用例
    # 生成测试报告的路径
    run_case(all_case)      # 2执行用例
    # 获取最新的测试报告文件
    report_path = os.path.join(cur_path, "report")  # 用例文件夹
    report_file = get_report_file(report_path)  # 3获取最新的测试报告
    #邮箱配置
    from config import readConfig
    sender = readConfig.sender
    psw = readConfig.psw
    smtp_server = readConfig.smtp_server
    port = readConfig.port
    receiver = readConfig.receiver
    send_mail(sender, psw, receiver, smtp_server, report_file, port)  # 4最后一步发送报告



