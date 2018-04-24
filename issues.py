#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import base64
import json
import oauth2 as oauth
import urllib
import pprint
from datetime import datetime
from tlslite.utils import keyfactory
import pymongo
import pprint
import pandas as pd
import numpy as np

jira_server_url = "http://bug.chenyee.com:8080/"
headers = {"Content-Type": "application/json"}
dept_info = {}
excluded_projects = ['ZPRJTE', 'TESTTOOLS', 'TEMP', 'TRANSTOOL', 'SCLUB', 'OA', 'DSHOP']


class SignatureMethod_RSA_SHA1(oauth.SignatureMethod):
    name = 'RSA-SHA1'

    def signing_base(self, request, consumer, token):
        if not hasattr(request, 'normalized_url') or request.normalized_url is None:
            raise ValueError("Base URL for request is not set.")

        sig = (
            oauth.escape(request.method),
            oauth.escape(request.normalized_url),
            oauth.escape(request.get_normalized_parameters()),
        )

        key = '%s&' % oauth.escape(consumer.secret)
        if token:
            key += oauth.escape(token.secret)
        raw = '&'.join(sig)
        return key, raw

    def sign(self, request, consumer, token):
        """Builds the base signature string."""
        key, raw = self.signing_base(request, consumer, token)

        with open('oauth_key/mykey.pem', 'r') as f:
            data = f.read()
        privateKeyString = data.strip()

        privatekey = keyfactory.parsePrivateKey(privateKeyString)
        signature = privatekey.hashAndSign(bytes(raw, 'utf-8'))

        return base64.b64encode(signature)


def get_depts(host='localhost'):
    client = pymongo.MongoClient(host)
    cydb = client.cy

    depts = cydb.dept.find(projection={'name': 1, 'dept': 1, 'group': 1, '_id': 0})

    for dept in depts:
        dept_info[dept['name']] = [dept['dept'].split('/')[1], dept['group']]


def get_client_handler():
    consumer_key = 'OauthKey'
    consumer_secret = 'dont_care'
    consumer = oauth.Consumer(consumer_key, consumer_secret)

    access_token = {'oauth_token': 'QHnpXBfHjXhYKfTMVDLMTyInQCefANgv', 'oauth_token_secret': 'k20W95a4ncIhSwBAcvlrXsIQn1IlZzcK'}
    accessToken = oauth.Token(access_token['oauth_token'], access_token[
                              'oauth_token_secret'])
    client = oauth.Client(consumer, accessToken)
    client.set_signature_method(SignatureMethod_RSA_SHA1())

    return client


def get_issues_from_jira(project_key=None, updated='50m', jira_webclient=None):
    search_url = jira_server_url + 'rest/api/2/search'
    maxResults = 4000  # jira.search.views.default.max
    jql = "issuetype = 故障"
    if project_key is not None:
        jql = "project = {0} && issuetype = 故障".format(project_key)

    if updated is not None:
        jql = jql + ' && updated >=-' + updated
    print(jql)
    search_key = {
        "jql": jql,
        "startAt": 0,
        "maxResults": maxResults,
        "expand": ['changelog'],
        "fields": [
            "summary",
            "components",
            "project",
            "reporter",
            "created",
            "priority",
            "customfield_10011",
            "customfield_10012",
            "customfield_10009",
            "assignee",
            "status",
            "updated",
            "resolution",
            "resolutiondate"
        ]
    }

    # get total issue
    resp, content = jira_webclient.request(method="POST", uri=search_url, headers=headers,
                                           body=bytes(json.dumps(search_key), 'utf-8'))
    issues_list = []
    if resp.status == 200:
        issues = json.loads(content.decode())
        issues_list.append(issues)
        total = issues['total']
        print("total issues is {0}".format(issues['total']))
        if total > maxResults:
            round = int(total/maxResults)
            for i in range(0, round):
                print("get more issues at round {0}".format(i))

                search_key['startAt'] = maxResults*(i+1)
                resp, content = jira_webclient.request(method="POST", uri=search_url, headers=headers, body=bytes(json.dumps(search_key), 'utf-8'))
                if resp.status == 200:
                    issues = json.loads(content.decode())
                    issues_list.append(issues)
                else:
                    print("error with {0}".format(resp.status))
    else:
        print("error with {0}".format(resp))

    return issues_list


def convert_changelog(changelog):
    changelog_mongodb = []

    for history in changelog['histories']:
        h = {}
        h['author'] = history['author']['name']
        h['date'] = datetime.strptime(history['created'], "%Y-%m-%dT%H:%M:%S.%f%z")

        h['items'] = []
        for item in history['items']:
            # just get status change historys
            if item['field'] == 'status':
                value = {'field': item['field'], 'from': item['fromString'], 'to': item['toString']}
                h['items'].append(value)
            elif item['field'] == 'assignee':
                value = {'field': item['field'], 'from': item['from'], 'to': item['to']}
                h['items'].append(value)

        if h['items'] != []:
            changelog_mongodb.append(h)

    return changelog_mongodb


def convert_data(issue):
    issue_for_mongodb = {}
    issue_for_mongodb['key'] = issue['key']
    issue_for_mongodb['project'] = issue['fields']['project']['key']
    issue_for_mongodb['component'] = issue['fields']['components'][0]['name']
    if issue['fields']['assignee'] is not None:
        try:
            issue_for_mongodb['assignee'] = {'name': issue['fields']['assignee']['name'],
                                             'displayName': issue['fields']['assignee']['displayName'],
                                             'dept': dept_info[issue['fields']['assignee']['name']][0],
                                             'group': dept_info[issue['fields']['assignee']['name']][1]}
        except KeyError:
            issue_for_mongodb['assignee'] = {'name': issue['fields']['assignee']['name'],
                                             'displayName': issue['fields']['assignee']['displayName'],
                                             'dept': '3rd', 'group': '3rd'}
            print("not found key: {}".format(issue['fields']['assignee']['name']))

    else:
        issue_for_mongodb['assignee'] = None
    issue_for_mongodb['reporter'] = {'name': issue['fields']['reporter']['name'],
                                     'displayName': issue['fields']['reporter']['displayName']}
    issue_for_mongodb['summary'] = issue['fields']['summary']
    issue_for_mongodb['status'] = issue['fields']['status']['name']
    issue_for_mongodb['priority'] = issue['fields']['priority']['name']
    issue_for_mongodb['probability'] = issue['fields']['customfield_10011']['value']
    issue_for_mongodb['severity'] = issue['fields']['customfield_10009']['value']
    issue_for_mongodb['phenomenon'] = issue['fields']['customfield_10012']['value']

    if issue['fields']['resolution'] != None:
        issue_for_mongodb['resolution'] = {'how': issue['fields']['resolution']['name'],
                                           'when': datetime.strptime(issue['fields']['resolutiondate'], "%Y-%m-%dT%H:%M:%S.%f%z")}
    else:
        issue_for_mongodb['resolution'] = {'how': None, 'when': None}

    issue_for_mongodb['created_time'] = datetime.strptime(issue['fields']['created'], "%Y-%m-%dT%H:%M:%S.%f%z")
    issue_for_mongodb['updated_time'] = datetime.strptime(issue['fields']['updated'], "%Y-%m-%dT%H:%M:%S.%f%z")

    if issue_for_mongodb['status'] == 'Closed':
        issue_for_mongodb['closed_time'] = issue_for_mongodb['updated_time']
    else:
        issue_for_mongodb['closed_time'] = None

    issue_for_mongodb['change_logs'] = convert_changelog(issue['changelog'])

    return issue_for_mongodb


def convert_issues(issues_list):
    # join with department
    issue_list = []
    for issues in issues_list:
        for issue in issues['issues']:
            if issue['fields']['project']['key'] not in excluded_projects:
                issue_list.append(convert_data(issue))
            else:
                print('skip ' + issue['fields']['project']['key'])

    print('finish converting')
    return issue_list


def create_mongodb(issue_list, host='localhost'):
    client = pymongo.MongoClient(host)
    cydb = client.cy
    cydb.issues.drop()
    cydb.issues.insert_many(issue_list)


def update_mongodb(issue_list, host='localhost'):
    client = pymongo.MongoClient(host)
    cydb = client.cy
    for issue in issue_list:
        result = cydb.issues.replace_one({'key': issue['key']}, issue, True)
        if result.raw_result['ok'] != 1.0:
            print("error: key is " + issue['key'])


def main(orig_args):
    mongodb_host = '18.8.8.209'
    jira_webclient = get_client_handler()
    get_depts(mongodb_host)
    issues_list = get_issues_from_jira(jira_webclient=jira_webclient,updated='1h')
    #issues_list = get_issues_from_jira(jira_webclient=jira_webclient,updated=None)
    print('sync with jira....' + str(datetime.now()))
    issue_list = convert_issues(issues_list)

    #create_mongodb(issue_list, mongodb_host)
    update_mongodb(issue_list, mongodb_host)

if __name__ == '__main__':
    main(sys.argv[1:])
