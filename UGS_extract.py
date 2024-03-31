"""
    Developed by AS on 31.03.2024
    
    The script works with UGS export folders and USG tunning parameters files named TuningParameter.sqlite.
    
    Each controller in UGS export folders contains two files: Tags_1.csv and Signals_1.csv, which have 
    common column named 'TagName'.
    Tags_1.csv contain tags, and 'TagName' column  values are unique across all Tags_1.csv files.
    Signals_1.csv contain signals, and there may be more than one row corresponing to a unique row in Tags_1.csv, 
    so it's kind of one-to-many relationship. However for most controllers (except RTU01/RTU02 with NPAS library) there 
    is only one row in Signals_1.csv with not empty IOAddress value corresponing to a row in Tags_1.csv, so it can 
    be processed as one-to-one relationship.
    
    Simplified algorithms for each cotroller:
        - process Tags_1.csv, remember every row (the only columns we need including TagName)
        - process Signals_1.csv, to each row remembered in the previous step add required columns from Signals_1.csv
          with the same TagName, but only for rows where IOAddress is not empty.
        - For each tag saved in the previous steps, if InstrumentType of a tag is USD-F64, search USG tunning parameters 
          files using TagName parameter and append LL, PL, PH, HH paramaters for such tags.

"""

import os
import sqlite3
import csv
import ctypes
import openpyxl as xl
from openpyxl.styles import Font
from datetime import datetime

def message_box(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)

os.system('cls')
print("******  start of the script  *******")

con1 = sqlite3.connect("TP/BCVU1151/TuningParameter.sqlite")
con2 = sqlite3.connect("TP/BCVU1152/TuningParameter.sqlite")
cur1 = con1.cursor()
cur2 = con2.cursor()

tags = {}
title = ["UGS_Name", "Controller Name", "TagName", "TagComment", "InstrumentType", "SL", "SH", "EU_PV",
         "IOAddress", "AccessType", "ScanName", "ConversionTypeName", "LL", "PL", "PH", "HH"]

for root, dirs, files in os.walk("BCVU export"):  
    UGS_name = root[root.find('\\')+1:root.rfind('\\')]
    controller_name = root[root.rfind("\\") + 1:]
    if files == ['Signals_1.csv', 'Tags_1.csv'] and controller_name[:4] != 'RTU0':  # exclude GTT controllers, cause they have multiple signals for one tag
        # tags are unique, so we can use them as a dict. key
        with open(os.path.join(root, "Tags_1.csv"), newline='', encoding='UTF-8') as csvfile:
            next(csvfile)
            for row in csv.reader(csvfile, delimiter=','):
                xl_row = []
                xl_row.append(UGS_name)           # UGS name
                xl_row.append(controller_name)    # Controller Name
                xl_row.append(row[0])             # TagName
                xl_row.append(row[3])             # TagComment
                xl_row.append(row[4])             # InstrumentType
                xl_row.append(row[15])            # SL
                xl_row.append(row[14])            # SH
                xl_row.append(row[16])            # EU_PV
                tags[row[0]] = xl_row
        with open(os.path.join(root, "Signals_1.csv"), newline='', encoding='UTF-8') as csvfile:
            next(csvfile)
            for row in csv.reader(csvfile, delimiter=','):
                if row[5] != '':                                # if IOAddress is not empty
                    TagName = row[0]                            # TagName
                    tags[TagName].append(row[5])                # IOAddress
                    tags[TagName].append(row[7])                # AccessType
                    tags[TagName].append(row[8])                # ScanName
                    tags[TagName].append(row[10])               # ConversionTypeName
                    
        # break # 1 iteration for test purposes

wb = xl.Workbook()  # create new workbook
ws = wb.active      
ws.title = "UGS signals"
ws.append(title)

for tag, row in tags.items():
    if row[4] == 'USD-F64':     # InstrumentType
        sql_select = f""" 
                        SELECT d.DataItemName, d.Value FROM Tag t, DataItem d 
                        WHERE TagName = '{tag}' and d.TagID = t.TagID and d.DataItemName in ('LL', 'PL', 'PH', 'HH') 
                        ORDER BY 1 DESC
                    """
        setpoints = None
        if (row[0] == 'BCVU1151'):
            res = cur1.execute(sql_select)
            setpoints = cur1.fetchall()
        if (row[0] == 'BCVU1152'):
            res = cur2.execute(sql_select)
            setpoints = cur2.fetchall()
        if setpoints is not None:
            if len(setpoints) > 0:
                sp_cols = [x[1] for x in setpoints]
                row += sp_cols
    ws.append(row)

for cell in ws['A1':'X1']:
    for x in cell:
        x.font = Font(bold=True)

dt = datetime.now().strftime(r"%Y-%m-%d")
try:
    wb.save(f"UGS Extract_{dt}.xlsx")
except  PermissionError as e:
    message_box("Error", str(e), 0)

con1.close()
con2.close()

print("******    end of the script    *******")