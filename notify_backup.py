#!/usr/bin/env python
# -*- coding:utf-8 -*-
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


class Mailer(object):

    def __init__(self, maillist, mailtitle, mailcontent, mailattach):
        self.mail_list = maillist
        self.mail_title = mailtitle
        self.mail_content = mailcontent
        self.mail_attachment = mailattach

        self.mail_host = "smtphm.qiye.163.com"
        self.mail_user = "scmjira@chenyee.com"
        self.mail_pass = "chenyee@2017"
        self.mail_postfix = "chenyee.com"

    def sendMail(self):

        me = self.mail_user #+ "<" + self.mail_user + ">"
        msg = MIMEMultipart()
        msg['Subject'] = self.mail_title
        msg['From'] = me
        msg['To'] = ";".join(self.mail_list)

        puretext = MIMEText('<h3>'+self.mail_content+'</h3>', 'html', 'utf-8')
        #puretext = MIMEText('纯文本内容'+self.mail_content)
        msg.attach(puretext)

        # jpg类型的附件
        #jpgpart = MIMEApplication(open('/home/mypan/1949777163775279642.jpg', 'rb').read())
        #jpgpart.add_header('Content-Disposition', 'attachment', filename='beauty.jpg')
        # msg.attach(jpgpart)

        # 首先是xlsx类型的附件
        #xlsxpart = MIMEApplication(open('test.xlsx', 'rb').read())
        #xlsxpart.add_header('Content-Disposition', 'attachment', filename='test.xlsx')
        # msg.attach(xlsxpart)

        # mp3类型的附件
        #mp3part = MIMEApplication(open('kenny.mp3', 'rb').read())
        #mp3part.add_header('Content-Disposition', 'attachment', filename='benny.mp3')
        # msg.attach(mp3part)

        # pdf类型附件
        part = MIMEApplication(open(self.mail_attachment, 'rb').read())
        part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(self.mail_attachment))
        msg.attach(part)

        try:
            s = smtplib.SMTP()  # 创建邮件服务器对象
            s.connect(self.mail_host)  # 连接到指定的smtp服务器。参数分别表示smpt主机和端口
            s.login(self.mail_user, self.mail_pass)  # 登录到你邮箱
            s.sendmail(me, self.mail_list, msg.as_string())  # 发送内容
            s.close()
            return True
        except Exception as e:
            print(str(e))
            return False


if __name__ == '__main__':
    # send list
    mailto_list = ["lugf@chenyee.com", "luusuu@126.com"]
    mail_title = 'Hey subject'
    mail_content = 'Hey this is content'
    mm = Mailer(mailto_list, mail_title, mail_content, '/Users/mmuunn/Documents/Works/issue_report/pp.pdf')
    res = mm.sendMail()
    print(res)
