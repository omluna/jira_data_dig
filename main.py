#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import pymongo
import datetime
import pandas as pd
import numpy as np
import shutil
from issue_report import *
from notify import *
from webdriver import *
import logging

logging.basicConfig(format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %H:%M:%S', 
        level=logging.INFO,
        filename='/home/lugf/reports/gen_reports.log')

logging.info("start generate reports")

work_dir = '/home/lugf/reports/'
employee_bugs_dir = work_dir + 'employee_bugs/'
employee_eff_dir = work_dir + 'employee_eff/'
bug_trend_dir = work_dir + 'bug_trend/'
dept_bug_stat_dir = work_dir + 'dept_stat_bugs/'

os.makedirs(dept_bug_stat_dir, exist_ok=True)
os.makedirs(employee_bugs_dir, exist_ok=True)
os.makedirs(bug_trend_dir, exist_ok=True)
os.makedirs(employee_eff_dir, exist_ok=True)

#project_list = ['SW17W16', 'CSW1702', 'CSW1705']
project_list = ['SW17W16', 'CSW1702', 'CSW1705', 'CSW1703']

client = pymongo.MongoClient('18.8.8.209')
cydb = client.cy


# pipeline = [ {'$project':{'_id':0, 'key':'$key', 'project':{ '$substr': [ "$project", 0, 7 ] }, 'assignee':'$assignee.displayName', 
#                           'dept':'$assignee.dept', 'group':'$assignee.group', 
#                           'status':'$status','created':{'$add': [ "$created_time", 8*60*60000 ]}, 'last_updated':{'$add': [ "$updated_time", 8*60*60000 ]},
#                           'change_logs':'$change_logs', 'resolved':{'$add': [ "$resolution.when", 8*60*60000 ]}, 
#                           'component':'$component','priority':'$priority','probability':'$probability', 'phenomenon':'$phenomenon'}},
#             {"$sort":{'assignee':1}}]


pipeline = [ {'$project':{'_id':0, 'key':'$key', 'project':{ '$substr': [ "$project", 0, 7 ] }, 'assignee_id':'$assignee.name', 'assignee':'$assignee.displayName', 
                          'dept':'$assignee.dept', 'group':'$assignee.group', 
                          'status':'$status','created':{'$add': [ "$created_time", 8*60*60000 ]}, 'last_updated':{'$add': [ "$updated_time", 8*60*60000 ]},
                          'change_logs':'$change_logs', 'resolved':{'$add': [ "$resolution.when", 8*60*60000 ]}, 
                          'component':'$component','priority':'$priority','probability':'$probability', 'phenomenon':'$phenomenon'}},
            {"$sort":{'created':1}}]

issues = pd.DataFrame(list(cydb.issues_new.aggregate(pipeline)))

now = datetime.datetime.today()
mm = Mailer()

def send_bug_trend():
    #BUG 趋势图
    report_dir = bug_trend_dir + str(now.date())
    os.makedirs(report_dir, exist_ok=True)
    os.chdir(report_dir)
    issue_p1 = issues[issues['priority'].isin(['P1-Highest'])]

    file_list = ['bugtrend.png']
    bug_trend(issue_p1, title='BUG趋势(P1)')

    for project in project_list:
        if len(issue_p1[issue_p1['project'] == project]) != 0:
            print("bug trend for " + project)
            bug_trend(issue_p1[issue_p1['project'] == project], title='{} BUG趋势(P1)'.format(project), image_filename=project+'_bugtrend')
            file_list.append(project+'_bugtrend.png')

    depts = {'sw': '平台及客户软件部', 'bsp': '驱动部', 'cam': '影像部'}
    for dept in iter(depts):
        if len(issue_p1[issue_p1['dept'] == depts[dept]]) != 0:
            print("bug trend for " + depts[dept])
            bug_trend(issue_p1[issue_p1['dept'] == depts[dept]], title='{} BUG趋势(P1)'.format(depts[dept]), image_filename=dept+'_bugtrend')
            file_list.append(dept+'_bugtrend.png')

    ppt_filename = str(now.date())+'_bugtrend.ppt'
    download_file(report_dir)
    make_pptx(file_list, ppt_filename)
    attachment=os.path.abspath(ppt_filename)
    mailto_list = ["lugf@chenyee.com", "gaolan@chenyee.com", "zhanght@chenyee.com", "spm@chenyee.com", 
            'yangyx@chenyee.com', 'liuxinhua@chenyee.com', 'luohui@chenyee.com', 'wangjz@chenyee.com']
    mail_title = 'P1 BUG趋势图-' + str(now.date())
    mail_content = '截至上周的P1BUG趋势图，见附件'
    #mailto_list = ['lugf@chenyee.com']
    if mm.sendemail('scmjira@chenyee.com', mailto_list, mail_title, mail_content, attachment=attachment) == False:
        logging.info('resending....')
        print('resend.....')
        mm.sendemail('scmjira@chenyee.com', mailto_list, mail_title, mail_content, attachment=attachment)
    logging.info('done sent bug trend....')
    print('-----done sent bug trend------')

# 每个部门每人BUG分布图，周期每周
def send_bug_employee_week():
    report_dir = dept_bug_stat_dir + str(now.date())
    os.makedirs(report_dir, exist_ok=True)
    os.chdir(report_dir)
    issue_p1 = issues[issues['priority'].isin(['P1-Highest'])]

    depts = {'sw': '平台及客户软件部', 'bsp': '驱动部', 'cam': '影像部'}
    leader_mail = {'sw':'luohui@chenyee.com', 'bsp':'yangyx@chenyee.com', 'cam':'zhanght@chenyee.com'}
    for dept in iter(depts):
        width = 800
        if len(issue_p1[issue_p1['dept'] == depts[dept]]) != 0:
            print("bug stat for " + depts[dept])
            if dept == 'sw':
                width=1200
            bug_employee_week(issue_p1, depts[dept], image_filename=dept, width=width)

    download_file(report_dir)

    for dept in iter(depts):
        mail_title = '{}成员P1BUG情况({})'.format(depts[dept], str(now.date()))
        mail_content = ' '
        mailto_list = [leader_mail[dept]]
        if dept == 'bsp':
            mailto_list.append('liuxinhua@chenyee.com')

        image_list = [os.path.abspath(dept + '.png')]
        mailcc = ['spm@chenyee.com', 'lugf@chenyee.com']

        print(dept + " " + str(mailto_list) + "\n" + str(image_list) + " mailcc:" + str(mailcc))
        #mailto_list = ['lugf@chenyee.com']
        #mailcc = ['lmmsuu@163.com','']
        if mm.sendemail('scmjira@chenyee.com', mailto_list, mail_title, mail_content, listCc=mailcc, listImagePath=image_list) == False:
            print('resend.....')
            mm.sendemail('scmjira@chenyee.com', mailto_list, mail_title, mail_content, listCc=mailcc, listImagePath=image_list)
        print('---finish send employee info ------')


# 每组BUG分布图，周期每天
def send_bug_group():
    dept_info = cydb.dept.find(projection={'name': 1, 'email': 1, 'dept': 1, 'group': 1, '_id': 0})
    depts = pd.DataFrame(list(dept_info))
    emails = dict(list(depts.groupby(depts['group'])['email']))

    group_list = {'xh':'晓慧组',  'xy':'小叶组', 'pp':'盼盼组', 'gms':'GMS组', 'xt':'小甜组', 'cam':'影像部', 'bsp':'驱动部', }
    leader_list = {'xh':'luohui@chenyee.com',  'xy':'luohui@chenyee.com', 'pp':'luohui@chenyee.com', 'gms':'luohui@chenyee.com', 
            'xt':'luohui@chenyee.com', 'cam':'', 'bsp':'', }
    report_dir = employee_bugs_dir + str(now.date())
    os.makedirs(report_dir, exist_ok=True)
    os.chdir(report_dir)
    unresolved_issue = issues[(issues.status.isin(['Open', 'Reopened', 'In Progress', 'Assigned']) & issues.project.isin(project_list))]
    for group in iter(group_list):
        bug_count_employee(unresolved_issue, group_list[group], project_list, filename=group)
    download_file(report_dir)

    for group in iter(group_list):
        mail_title = '{}未解BUG统计({})'.format(group_list[group], str(now.date()))
        mail_content = ' '
        if group == 'cam':
            mailto_list = ['image@chenyee.com']
        elif group == 'bsp':
            mailto_list = ['bsp@chenyee.com']
        else:
            mailto_list = list(emails[group_list[group]])
        image_list = [os.path.abspath(group + '.png')]
        mailcc = ['spm@chenyee.com', 'lugf@chenyee.com', leader_list[group]]

        print(group + " " + str(mailto_list) + "\n" + str(image_list) + " mailcc:" + str(mailcc))
        #mailto_list = ['lugf@chenyee.com']
        #mailcc = ['lmmsuu@163.com','']
        if mm.sendemail('scmjira@chenyee.com', mailto_list, mail_title, mail_content, listCc=mailcc, listImagePath=image_list) == False:
            print('resend.....')
            mm.sendemail('scmjira@chenyee.com', mailto_list, mail_title, mail_content, listCc=mailcc, listImagePath=image_list)
        print('---finish send group info ------')


def send_bug_group_eff():
    dept_info = cydb.dept.find(projection={'name': 1, 'email': 1, 'displayName':1, 'dept': 1, 'group':1, '_id': 0})
    depts = pd.DataFrame(list(dept_info))
    emails = dict(list(depts.groupby(depts['group'])['email']))

    group_list = {'xh':'晓慧组',  'xy':'小叶组', 'pp':'盼盼组', 'gms':'GMS组', 'xt':'小甜组', 'cam':'影像部', 'bsp':'驱动部', }
    leader_list = {'xh':'luohui@chenyee.com',  'xy':'luohui@chenyee.com', 'pp':'luohui@chenyee.com', 'gms':'luohui@chenyee.com', 
            'xt':'luohui@chenyee.com', 'cam':'', 'bsp':'', }
    report_dir = employee_eff_dir + str(now.date())
    os.makedirs(report_dir, exist_ok=True)
    os.chdir(report_dir)

    issues_p1_p2 = issues[issues['priority'].isin(['P1-Highest', 'P2-High'])]
    anchor_time_list = []
    for i, issue in issues_p1_p2.iterrows():
        anchor_time_list += parse_changelog(issue)
    
    df = pd.DataFrame(anchor_time_list)
    anchro_time_df = pd.merge(df, depts, left_on='who', right_on='name')

    for group in iter(group_list):
        group_member = anchro_time_df[anchro_time_df.group == group_list[group]].displayName.unique()
        for i, name in enumerate(group_member):
            stop_time_count_bubble(anchro_time_df, name, filename=group+str(i))

    download_file(report_dir)

    for group in iter(group_list):
        group_member = anchro_time_df[anchro_time_df.group == group_list[group]].displayName.unique()
        mail_title = '{}近两周BUG流转效率统计({})'.format(group_list[group], str(now.date()))
        mail_content = ' '
        if group == 'cam':
            mailto_list = ['image@chenyee.com']
        elif group == 'bsp':
            mailto_list = ['bsp@chenyee.com']
        else:
            mailto_list = list(emails[group_list[group]])
        image_list = [os.path.abspath('{}{}.png'.format(group,i)) for i in range(0, len(group_member))]
        mailcc = ['spm@chenyee.com', 'lugf@chenyee.com', leader_list[group]]

        print(group + " " + str(mailto_list) + "\n" + str(image_list) + " mailcc:" + str(mailcc))
        #mailto_list = ['lugf@chenyee.com']
        #mailcc = ['lmmsuu@163.com','']
        if mm.sendemail('scmjira@chenyee.com', mailto_list, mail_title, mail_content, listCc=mailcc, listImagePath=image_list) == False:
            print('resend.....')
            mm.sendemail('scmjira@chenyee.com', mailto_list, mail_title, mail_content, listCc=mailcc, listImagePath=image_list)
        print('---finish send group efficient info ------')


def main(orig_args):
    print("generate report for " + orig_args[0])
    if orig_args[0] == 'bugtrend':
        send_bug_trend()

    if orig_args[0] == 'day_bug_group':
        send_bug_group()

    if orig_args[0] == 'week_bug_employee':
        send_bug_employee_week()

    if orig_args[0] == 'group_eff':
        send_bug_group_eff()


if __name__ == '__main__':
    main(sys.argv[1:])

