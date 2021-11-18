# -*- coding: utf-8 -*-
"""
2021
@author: wuffalo
"""

import pandas as pd
#import xlsxwriter #included in pandas
import os
from datetime import datetime as dt, timedelta
import glob

def format_sheet(X):
    X = X+1
    worksheet.set_column('A:A',13)
    worksheet.set_column('B:B',45)
    worksheet.set_column('C:C',5)
    worksheet.set_column('D:D',7)
    worksheet.set_column('E:E',19)
    worksheet.set_column('F:F',17)
    worksheet.set_column('G:G',19)
    worksheet.set_column('H:H',11)
    worksheet.set_column('I:I',7,format6)
    worksheet.set_column('J:J',28)
    worksheet.set_column('K:K',14,format5)
    worksheet.conditional_format('K2:K'+str(X), {'type': 'duplicate',
                                        'format': format4})
    worksheet.conditional_format('E2:E'+str(X), {
        'type': 'date',
        'criteria': 'less than or equal to',
        'value': (ctime-timedelta(1)),
        'format': format1
        })
    worksheet.conditional_format('E2:E'+str(X), {
        'type': 'date',
        'criteria': 'between',
        'minimum': ctime-timedelta(11/12),
        'maximum': ctime-timedelta(1),
        'format': format2
        })
    worksheet.conditional_format('E2:E'+str(X), {
        'type': 'date',
        'criteria': 'between',
        'minimum': ctime-timedelta(4/5),
        'maximum': ctime-timedelta(11/12),
        'format': format3
        })

output_directory = "/mnt/shared-drive/04 - SLC Operations/24Hour Report/"
output_file_name = "24Hour.xlsx"
path_to_output = output_directory+output_file_name

def empty_folder(dir,name,ext): #directory to remove old files, search term and file extension are optional
    old_files = glob.glob(dir+'*'+name+'*'+ext)
    for f in old_files:
        try:
            os.remove(f)
        except:
            pass

if os.path.exists(path_to_output):
    if os.path.exists(output_directory+'~$'+output_file_name):
        empty_folder(output_directory,'24Hour','.xlsx')
        path_to_output = output_directory+os.path.splitext(output_file_name)[0]+str(dt.now().strftime(' %b%d %Hh%Mm '))+".xlsx"
    else:
        empty_folder(output_directory,'24Hour','.xlsx')

ctime = dt.now()

path_to_sharedSOS = '/mnt/shared-drive/Operations/Data/Shipment Order Summary (PICK ZONE).csv'

file_time_best = os.path.getctime(path_to_sharedSOS)

#Use best SOS for program
update_time = dt.fromtimestamp(file_time_best).strftime('%m/%d/%Y %H:%M')
df = pd.read_csv(path_to_sharedSOS, parse_dates=[11,19], infer_datetime_format=True)

#columns to delete - INITIAL PASS
df = df.drop(columns=['ORDERKEY','SO','SS','STORERKEY','INCOTERMS','ORDERDATE','ACTUALSHIPDATE','DAYSPASTDUE',
                'PASTDUE','ORDERVALUE','TOTALSHIPPED','EXCEP','STOP','PSI_FLAG','SUSR5','INTERNATIONALFLAG',
                'BILLING','LOADEDTIME','UDFVALUE1'])

#rename remaining columns
df = df.rename(columns={'EXTERNORDERKEY':'SO-SS','C_COMPANY':'Customer','ADDDATE':'Add Date','STATUSDESCR':'Status',
                        'TOTALORDERED':'QTY','SVCLVL':'Carrier','EXTERNALLOADID':'Load ID','EDITDATE':'Last Edit',
                        'C_STATE':'State','C_COUNTRY':'Country','Textbox6':'TIS'})

writer = pd.ExcelWriter(path_to_output, engine='xlsxwriter', options={'strings_to_numbers': True})
workbook = writer.book

# Light red fill with dark red text.
format1 = workbook.add_format({'bg_color':   '#FFC7CE',
                               'font_color': '#9C0006'})
# orange fill with dark orange text.
format2 = workbook.add_format({'bg_color':   '#ffcc99',
                               'font_color': '#804000'})
# yellow fill with dark yellow text.
format3 = workbook.add_format({'bg_color':    '#ffeb99',
                                'font_color':   '#806600'})
# Green fill with dark green text.
format4 = workbook.add_format({'bg_color':   '#C6EFCE',
                               'font_color': '#006100'})
format5 = workbook.add_format({'num_format': '#'})
format6 = workbook.add_format({'num_format': '#,##0'})

#CREATE QUERIES TO REMOVE
remove_rtv = df['TYPEDESCR'] == 'RTV Move'
remove_NS = df['Status'] == 'Not Started'
remove_Lo = df['Status'] == 'Loaded'

df_loaded = df[df['Status'] == 'Loaded']

df.drop(df[remove_rtv|remove_NS|remove_Lo].index, inplace=True)

# create column that floors Add Date by hour
df['Add Hour'] = df['Add Date'].dt.floor('1H')

# sort dataframes
df.sort_values(by=['Add Hour','Status','Carrier'], inplace=True)
df_loaded.sort_values(by=['Carrier']) # cannot use inplace because copy of df error

# drop columns
df = df.drop(columns=['TYPEDESCR','CUSTID','PROMISEDATE','Add Hour','ROUTE'])
df_loaded = df_loaded.drop(columns=['CUSTID','PROMISEDATE','Status','ROUTE'])

# calculate lengths of dataframes
main_length = len(df.index)
loaded_length = df_loaded.TIS.count()

# create and format sheet of most normal orders
df.to_excel(writer, sheet_name='24Hour', index=False)
worksheet = writer.sheets['24Hour']
worksheet.write('M1',"Last Update at: "+str(update_time))
format_sheet(main_length)

# create and format sheet of Loaded orders
df_loaded.to_excel(writer, sheet_name='Loaded', index=False)
worksheet = writer.sheets['Loaded']
format_sheet(loaded_length)

writer.save()