#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import pymongo
import pprint
import pandas as pd
import numpy as np
import plotly.offline as py
from issue_report import *
from notify import *
from webdriver import *
import shutil

py.init_notebook_mode(connected=True)

#remove temp_station
temp_station = '/home/lugf/py3/issue_report/temp_station/'
shutil.rmtree(temp_station)
os.mkdir(temp_station)


client = pymongo.MongoClient('18.8.8.209')
cydb = client.cy

pipeline = [{'$project': {'key': '$key', 'project': {'$substr': ["$project", 0, 7]}, 'assignee':'$assignee.displayName', 'dept':'$assignee.dept', 'status':'$status', 'created':'$created_time', 'change_logs':'$change_logs', 'resolved':'$resolution.when', 'component':'$component', 'priority':'$priority', 'probability':'$probability', 'phenomenon':'$phenomenon'}}, {"$sort": {'assignee': 1}}]

issues = pd.DataFrame(list(cydb.issues.aggregate(pipeline)))
del issues['_id']

issue_pd = issues[issues['priority'].isin(['P1-Highest'])]
issue_pd.count()['key']


project_list = ['SW17W16', 'CSW1702', 'CSW1705', 'CSW1703A']
bug_trend(issue_pd, title='BUG趋势(P1)')
file_list = ['bugtrend.png']

for project in project_list:
    if len(issue_pd[issue_pd['project'] == project]) != 0:
        bug_trend(issue_pd[issue_pd['project'] == project], title='{} BUG趋势(P1)'.format(project), image_filename=project+'_bugtrend')
        file_list.append(project+'_bugtrend.png')

depts = {'sw':'平台及客户软件部', 'bsp':'驱动部', 'cam':'影像部'}
for dept in iter(depts):
    if len(issue_pd[issue_pd['dept'] == depts[dept]]) != 0:
        bug_trend(issue_pd[issue_pd['dept'] == depts[dept]], title='{} BUG趋势(P1)'.format(depts[dept]), image_filename=dept+'_bugtrend')
        file_list.append(dept+'_bugtrend.png')

print(file_list)

ppt_filename = str(datetime.date.today())+'_bugtrend.ppt'
download_file(temp_station)
make_pptx(file_list, ppt_filename)


#mailto_list = ["lugf@chenyee.com", "luusuu@126.com"] #, "spm@chenyee.com", 'yangyx@chenyee.com', 'luohui@chenyee.com', 'wangjz@chenyee.com']
mailto_list = ["lugf@chenyee.com", "spm@chenyee.com", 'yangyx@chenyee.com', 'luohui@chenyee.com', 'wangjz@chenyee.com']
mail_title = 'P1 BUG趋势图'
mail_content = '先送上一份截止到上周的BUG趋势图，其他稍后奉上 :)'
mm = Mailer(mailto_list, mail_title, mail_content, temp_station+ppt_filename)
res = mm.sendMail()
print(res)
