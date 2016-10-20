#coding=utf-8
import ConfigParser
import smtplib
from email.mime.text import MIMEText

"""
此类的作用是生成测试报告(返回报告地址，文件打开的地址)
"""
from Share import HTMLTestRunner


def Report(FilePath, ReportTitle="Test Report", ReportDescription=None):
    fp=file(FilePath, 'wb')
    return HTMLTestRunner.HTMLTestRunner(stream=fp,title=ReportTitle, description=ReportDescription), fp


#发生邮件
"""
测试完成后或发生异常时
"""
#邮件用户名，密码，收件从，主题，内容，文件名，传送方式0表示文字 1表示：html
def SenEmail(user, password,receive,subject,details,filename,datatype=0):
    if datatype ==1:
        paper=open(filename,'r')
        details=paper.read()
        paper.close()
        msg = MIMEText(details,'html','utf-8')
    else:
        msg = MIMEText(details)
    msg["Subject"] = subject
    msg["From"]    = user
    msg["To"]      = receive
    try:
        s = smtplib.SMTP_SSL("smtp.qq.com", 465)
        s.login(user, password)
        s.sendmail(user, receive, msg.as_string())
        s.quit()
        print "Success!"
    except smtplib.SMTPException,e:
        print "Falied,%s"%e

if __name__ =='__main__':
    _user='234975828@qq.com'
    _pwd='wgwjgnwipjarcbdh'
    _to='904263480@qq.com'
    textsubject='报告'
    details='详情'
    hl='G:\Python\CashBox\CashBox\TestResult\Result\Test_Report.html'
    SenEmail(_user, _pwd, _to,textsubject,details,hl,1)