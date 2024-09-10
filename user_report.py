"""
Title: Read a Snellius data file created by ./get_user_data.py and convert it into an Excel file
Author: Peter Vos
version: 0.8-alpha

Usage (Local): python3 user_report.py

Requires: openpyxl (pip install openpyxl)

(C) Peter J.M. Vos, VU Amsterdam, 2024. Licenced for use under the BSD 3 clause

Create a json with ./get_user_data.py first.
"""

import openpyxl
from datetime import datetime
import json

def create_excel(datafile):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'snellius_usage'
    
    with open(f"{datafile}", "r+") as fp:
        data = json.load(fp)

    workbook = openpyxl.Workbook()
    sheet = workbook.active

    sheet.append(['email', 'account', 'company', 'department', 'eduPersonAffiliation', 'title', 'displayName', 'cpu budget', 'cpu usage', 'gpu budget', 'gpu usage', 'projectspace budget', 'projectspace usage'])
    i = 2
    for email, userdata in data.items():
        projectspace = userdata.get('projectspace', {})
        CPU = userdata.get('CPU', {})
        GPU = userdata.get('GPU', {})
        sheet.append([
            email,
            userdata['account'],
            userdata.get('company', ''),
            userdata.get('department', ''),
            userdata.get('eduPersonAffiliation', ''),
            userdata.get('title', ''),
            userdata.get('displayName', ''),
            CPU.get('budget', 0),
            round(CPU.get('usage', 0)),
            GPU.get('budget', 0),
            round(GPU.get('usage', 0)),
            projectspace.get('budget', 0),
            projectspace.get('usage', 0)
        ])
        i += 1

    workbook.save(filename = datafile.replace('.json', '.xlsx'))


if __name__ == '__main__':
    create_excel("data/snellius_usersAD-20240905.json")
