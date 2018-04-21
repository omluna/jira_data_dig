#!/usr/bin/env python
# -*- coding:utf-8 -*-
import os
import smtplib

from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

class Mailer(object):

    def __init__(self):
        self.smtpserver = "smtphm.qiye.163.com"
        self.smtpuser = "scmjira@chenyee.com"
        self.smtppwd = "chenyee@2017"

    def generateAlternativeEmailMsgRoot(self, strFrom, listTo, listCc, strSubJect, strMsgText, strMsgHtml, listImagePath, attachment):
        # Create the root message and fill in the from, to, and subject headers
        msgRoot = MIMEMultipart('related')
        msgRoot['Subject'] = strSubJect
        msgRoot['From'] = strFrom
        msgRoot['To'] = ",".join(listTo)
        if listCc:
            msgRoot['Cc'] = ",".join(listCc)
        msgRoot.preamble = 'This is a multi-part message in MIME format.'

        # Encapsulate the plain and HTML versions of the message body in an
        # 'alternative' part, so message agents can decide which they want to display.
        msgAlternative = MIMEMultipart('alternative')
        msgRoot.attach(msgAlternative)

        msgContent = strMsgText.replace("\n", "<br>") if strMsgText else ""
        msgContent += "<br>" + strMsgHtml if strMsgHtml else ""

        # We reference the image in the IMG SRC attribute by the ID we give it below
        if listImagePath and len(listImagePath) > 0:
            msgHtmlImg = msgContent + "<br>"
            for imgcount in range(0, len(listImagePath)):
                msgHtmlImg += '<img src="cid:image{count}"><br>'.format(count=imgcount)

            msgText = MIMEText(msgHtmlImg, 'html')
            msgAlternative.attach(msgText)

            # This example assumes the image is in the current directory
            for i, imgpath in enumerate(listImagePath):
                print(imgpath)
                fp = open(imgpath, 'rb')
                msgImage = MIMEImage(fp.read())
                fp.close()

                # Define the image's ID as referenced above
                msgImage.add_header('Content-ID', '<image{count}>'.format(count=i))
                msgRoot.attach(msgImage)
        else:
            msgText = MIMEText(msgContent, 'html')
            msgAlternative.attach(msgText)

        if attachment:
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
            part = MIMEApplication(open(attachment, 'rb').read())
            part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(attachment))
            msgRoot.attach(part)

        return msgRoot

    # Send the email (this example assumes SMTP authentication is required)
    def sendemail(self, strFrom, listTo, strSubJect, strMsgText, strMsgHtml=None, listImagePath=None, listCc=None, attachment=None):
        msgRoot = self.generateAlternativeEmailMsgRoot(strFrom, listTo, listCc, strSubJect, strMsgText, strMsgHtml, listImagePath, attachment)

        try:
            self.smtp = smtplib.SMTP()
            self.smtp.connect(self.smtpserver)
            self.smtp.login(self.smtpuser, self.smtppwd)
            if listCc:
                listTo = listTo + listCc
            self.smtp.sendmail(strFrom, listTo, msgRoot.as_string())
            self.smtp.quit()
            print("Send mail success {0}".format(strSubJect))
            return True
        except Exception as e:
            print("ERROR:Send mail failed {0} with {1}".format(strSubJect, str(e)))
            #self.smtp.quit()
            return False


if __name__ == '__main__':
    # send list
    mailto_list = ["lugf@chenyee.com", "luusuu@126.com", 'zhangmx@chenyee.com']
    mail_title = 'Hey subject'
    mail_content = 'Hey this is content'
    os.chdir('/Users/mmuunn/Desktop/1')
    picFiles = [fn for fn in os.listdir() if fn.endswith('.png')]

    image_list = []
    for pic in picFiles:
        image_list.append(os.path.abspath(pic))

    print(image_list)

    mm = Mailer()
    mm.sendemail('lugf@chenyee.com', mailto_list, "text mail", "Hi it's Max, this is a test maill-----1", "<h2>test html content</h2>", attachment='/Users/mmuunn/Documents/Works/issue_report/pp.pdf', listImagePath=image_list)
