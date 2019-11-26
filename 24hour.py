# -*- coding: utf-8 -*-
"""
2019
@author: wuffalo
"""

import pandas as pd
import xlsxwriter
import os
import glob

def format_sheet():
    worksheet.set_column('A:A',13)
    worksheet.set_column('B:B',45)
    worksheet.set_column('C:C',5)
    worksheet.set_column('D:D',7)
    worksheet.set_column('E:E',20)
    worksheet.set_column('F:F',14)
    worksheet.set_column('G:G',21)
    worksheet.set_column('H:H',11)
    worksheet.set_column('I:I',5)
    worksheet.set_column('J:J',28)
    worksheet.set_column('K:K',14,format5)
    worksheet.conditional_format('K2:K4000', {'type': 'duplicate',
                                              'format': format3})

path_to_excel = "/mnt/shared-drive/05 - Office/OTS/Wolf/24Hour.xlsx"

if os.path.exists(path_to_excel):
  os.remove(path_to_excel)

list_of_files = glob.glob('/mnt/c/Users/WMINSKEY/Downloads/Shipment Order Summary -*.csv') # * means all if need specific format then *.csv
latest_file = max(list_of_files, key=os.path.getctime)
path_to_SOS = latest_file

df = pd.read_csv(path_to_SOS, parse_dates=[11,19], infer_datetime_format=True)

#columns to delete - INITIAL PASS
df = df.drop(columns=['ORDERKEY','SO','SS','STORERKEY','INCOTERMS','ORDERDATE','ACTUALSHIPDATE','DAYSPASTDUE',
                'PASTDUE','ORDERVALUE','TOTALSHIPPED','EXCEP','STOP','PSI_FLAG','UDFNOTES','INTERNATIONALFLAG',
                'BILLING','LOADEDTIME','UDFVALUE1'])

#rename remaining columns
df = df.rename(columns={'EXTERNORDERKEY':'SO-SS','C_COMPANY':'Customer','ADDDATE':'Add Date','STATUSDESCR':'Status',
                        'TOTALORDERED':'QTY','SVCLVL':'Carrier','EXTERNALLOADID':'Load ID','EDITDATE':'Last Edit',
                        'C_STATE':'State','C_COUNTRY':'Country','Textbox6':'TIS'})

writer = pd.ExcelWriter(path_to_excel, engine='xlsxwriter')
workbook = writer.book

# Light red fill with dark red text.
format1 = workbook.add_format({'bg_color':   '#FFC7CE',
                               'font_color': '#9C0006'})

# Light yellow fill with dark yellow text.
format2 = workbook.add_format({'bg_color':   '#FFEB9C',
                               'font_color': '#9C6500'})

# Green fill with dark green text.
format3 = workbook.add_format({'bg_color':   '#C6EFCE',
                               'font_color': '#006100'})

format5 = workbook.add_format({'num_format': '#'})

#CREATE QUERIES TO REMOVE
remove_rtv = df['TYPEDESCR'] == 'RTV Move'
remove_NS = df['Status'] == 'Not Started'
remove_Lo = df['Status'] == 'Loaded'

df.drop(df[remove_rtv|remove_NS|remove_Lo].index, inplace=True)

df['Add Hour'] = df['Add Date'].dt.floor('1H')

df.sort_values(by=['Add Hour','Status','Carrier'], inplace=True)

df = df.drop(columns=['TYPEDESCR','CUSTID','PROMISEDATE','Add Hour'])

df.to_excel(writer, sheet_name='24Hour', index=False)
worksheet = writer.sheets['24Hour']
format_sheet()

writer.save()