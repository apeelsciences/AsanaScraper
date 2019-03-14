# -*- coding: utf-8 -*-
import json
import csv
from pprint import pprint
import xlwt
import re
import sys

with open('tasks.json') as f:
    data = json.load(f)

workbook = xlwt.Workbook()
sheet = workbook.add_sheet("IT Support Request Data")

row = 0
for item in data["data"]:

    # retrieve task title
    sheet.write(row, 0, item['name'])

    # retrieve assigned tech, empty if none
    try:
        sheet.write(row, 1, item['assignee']['name'])
    except:
        sheet.write(row, 1, '')

    # custom fields
    custom = item['notes'].replace('\r', '').replace('\n', '')

    # parse and write user from 'notes' field

    pattern = re.compile(ur'(?<=Email::)(.*)(?=Subject::)')
    match = (pattern.findall(custom))
    user = ''.join(match)
    sheet.write(row, 2, user)

    # parse and write Type of IT Request from 'notes' field
    pattern = re.compile(ur'(?<=Request::)(.*)(?=Other::|Business|Description|Question:)')
    match = (pattern.findall(custom))
    requestType = ''.join(match)
    sheet.write(row, 3, requestType)

    #Date and time
    sheet.write(row, 4, item['created_at'])

    # parse and write Severity from 'notes' field
    pattern = re.compile(ur'(?<=Severity::)(.*)(?=Attach)')
    match = (pattern.findall(custom))
    severity = ''.join(match)
    sheet.write(row, 5, severity)

    # parse and write message body
    pattern = re.compile(ur'(?<=ther::|Case::|ssue::|tion::)(.*)(?=Severity::)')
    match = (pattern.findall(custom))
    body = ''.join(match)
    sheet.write(row, 6, body)

    row += 1



workbook.save('test.xls')