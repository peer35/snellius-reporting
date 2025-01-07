"""
Title: Read a Snellius Excel report file and enrich with AD info
Author: Peter Vos
version: 0.9

Usage (Local): python3 get_user_data.py

Requires: openpyxl (pip install openpyxl), ldap3 (pip install ldap3)

(C) Peter J.M. Vos, VU Amsterdam, 2024. Licenced for use under the BSD 3 clause

Make sure to set the AD credentials and the path to the synced reports in ./config.py
"""

from config import REPORT_PATH
import openpyxl
import json

def get_headings(sheet):
    headings = []
    for n in range(1, sheet.max_column):
        headings.append(sheet.cell(1, column=n).value)
    print(headings)
    return headings


def create_json_from_report(reportfile, ad_lookup=False):
    reportdate = reportfile.split(".")[1]
    workbook = openpyxl.load_workbook(f"{REPORT_PATH}/{reportfile}")
    sheet = workbook.active
    headings=get_headings(sheet)
    col_account = headings.index("Account")+1
    try:
        col_email = headings.index("email")+1
    except ValueError:
        col_email = False
    col_usage = headings.index("SrvUsage")+1
    col_budget = headings.index("Budget")+1
    col_trend = headings.index("trend")+1 # should be the last column before the monthly usage columns

    data = {}
    datafile = f"{REPORT_PATH}/{reportfile.replace('.xlsx','.json')}"
    data = {}
    
    for i in range(1, sheet.max_row + 1):
        print(i)
        code = sheet.cell(row=i, column=1).value
        description = sheet.cell(row=i, column=2).value
        #  use the "sub-heading"-rows in the decription column to set the type of budget
        if description == "Snellius VU CPU-compute 2024":
            budget_type = "CPU"
        elif description == "Snellius VU GPU-compute 2024":
            budget_type = "GPU"
        elif description == "Snellius VU project storage 2024":
            budget_type = "projectspace"
        account = sheet.cell(row=i, column=col_account).value
        if account.startswith("snel-vusr") and code.startswith("2307090"):
            if account not in data:
                data[account]={}
            budget = sheet.cell(row=i, column=col_budget).value
            print(account, budget)
            usage = sheet.cell(row=i, column=col_usage).value
            if col_email:
                email = sheet.cell(row=i, column=col_email).value
                if email not in data:
                    data[account]["email"] = email
            if budget_type in data[account]:
                data[account][budget_type]["budget"] += budget
                data[account][budget_type]["usage"] += usage
            else:
                data[account][budget_type] = {"budget": budget, "usage": usage}

    # dump the data in a json for now
    with open(datafile, "w") as fp:
        json.dump(data, fp)
    return datafile


if __name__ == "__main__":
    datafile=create_json_from_report(reportfile="2307090_24.20250106.xlsx")
