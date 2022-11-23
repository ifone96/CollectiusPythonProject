from ast import If
import datetime
import glob
import os
import shutil
from tkinter import HIDDEN
import uuid
from doctest import DocFileTest
from email.utils import format_datetime
from math import fabs
from operator import index
from pickle import NONE
import pandas as pd
import pyodbc
import xlsxwriter
from matplotlib.pyplot import axis
import zipfile
import numpy as np


data_file_folder = "C:\\Users\\wasin.k\\OneDrive - COLLECTIUS SYSTEMS PTE. LTD\\Ulite\\New Assignment\\" 
df = []
for file in os.listdir(data_file_folder):
    if file.endswith('.xlsx'):
        print('Loading file {0}...'.format(file))
        df.append(pd.read_excel(os.path.join(data_file_folder, file), sheet_name='follow_debt_report', header=11))
    if file.endswith('.xlsb'):
        print('Loading file {0}...'.format(file))
        df.append(pd.read_excel(os.path.join(data_file_folder, file), sheet_name='follow_debt_report', header=11))
      
      
# Len(df)
df_combine = pd.concat(df, axis=0)

df_combine = df_combine[[
                        'ผู้ซื้อ',
                        'ชื่อ-นามสกุล',
                        'Order No',
                        'ชื่อสินค้า',
                        'เบอร์โทร',
                        'สถานะการตามหนี้',
                        'โทรโดย',
                        'เวลาโทรล่าสุด',
                        'ประเภท',
                        'สถานะการติดตาม',
                        'การติดตามลูกค้า',
                        'การติดต่อผู้อ้างอิง',
                        'หมายเหตุการโทร',
                        'กำหนดชำระ',
                        'งวด',
                        'จำนวนวัน',
                        'ภาระหนี้ค้างชำระ(Over Due)(บาท)',
                        'รวมภาระหนี้คงค้าง',
                        'ค่าปรับ+ค่าทวงถาม(บาท)',
                        'รวมค่าปรับ+ค่าทวงถาม',
                        'ค่าปรับ(บาท)',
                        'ค่าทวงถาม(บาท)',
                        'ภาระหนี้ค้างชำระ + ค่าปรับ(บาท)',
                        'วันนัดชำระ',
                        'ผิดนัดชำระ',
                        'จ่ายแล้ว',
                        'ส่ง Outsource',
                        'วันที่ส่ง Outsource',
                        'Product Program',
                        'มหาวิทยาลัย',
                        'คณะ',
                        'ชั้นปี',
                        'บุคคลอ้างอิง(ชื่อ)',
                        'บุคคลอ้างอิง(โทร)',
                        'บุคคลอ้างอิง(สัมพันธ์)',
                        'บุคคลอ้างอิง2(ชื่อ)',
                        'บุคคลอ้างอิง2(โทร)',
                        'บุคคลอ้างอิง2(สัมพันธ์)',
                        'บุคคลอ้างอิง3(ชื่อ)',
                        'บุคคลอ้างอิง3(โทร)',
                        'บุคคลอ้างอิง3(สัมพันธ์)',
                        'เลขที่บัตรประชาชนใบเสร็จ',
                        'ที่อยู่',
                        'จังหวัด',
                        'เขต/อำเภอ',
                        'แขวง/ตำบล',
                        'รหัสไปรษณีย์',
                        'ที่อยู่(ปัจจุบัน)',
                        'จังหวัด(ปัจจุบัน)',
                        'เขต/อำเภอ(ปัจจุบัน)',
                        'แขวง/ตำบล(ปัจจุบัน)',
                        'รหัสไปรษณีย์(ปัจจุบัน)',
                        'สถานะ Fraud',
                        'วันเกิด',
                        'อีเมล',
                        'วันที่รับชำระก่อน',
                        'หัวข้อ Notice',
                        'เวลาออก Notice',
                        'เบอร์โทร(2)',
                        'เบอร์โทรศัพท์บ้าน',
                        'Line',
                        'Twitter',
                        'Facebook',
                        'Instragram',
                        'มีรายได้เสริม',
                        'เวลาที่สะดวกติดต่อ',
                        'Aging',
                         ]]

df_combine.insert(59, 'Write Off', '', True)
df_combine = df_combine[df_combine['รวมภาระหนี้คงค้าง'] > 0]

todaysdate_filename = str(
    datetime.datetime.now().strftime("CombineUliteFollowDebt'%Y%m%d")) + '.xlsx'
writer = pd.ExcelWriter(todaysdate_filename)

df_combine.to_excel(writer, index=False, sheet_name= 'CombineUliteFollowDebt')
# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['CombineUliteFollowDebt']


# Add some cell formats.
format = workbook.add_format({'num_format': '###########################'})
format1 = workbook.add_format({'num_format': '@'})
format2 = workbook.add_format({'num_format': '#############'})
format3 = workbook.add_format({'num_format': 'mm/dd/yyyy'})
formatphone = workbook.add_format({'num_format': '0000000000'})

header_format = workbook.add_format({'bold': True})

worksheet.set_row(0, None, header_format)
worksheet.set_column('A:A', 25, format)
worksheet.set_column('B:B', 25, format)
worksheet.set_column('C:C', 25, format)
worksheet.set_column('D:D', 25, format)
worksheet.set_column('E:E', 25, formatphone)
worksheet.set_column('F:F', 25, format)
worksheet.set_column('G:G', 25, format)
worksheet.set_column('H:H', 25, format)
worksheet.set_column('I:I', 25, format)
worksheet.set_column('J:J', 25, format)
worksheet.set_column('K:K', 25, format)
worksheet.set_column('L:L', 25, format)
worksheet.set_column('M:M', 25, format)
worksheet.set_column('N:N', 25, format)
worksheet.set_column('O:O', 25, format)
worksheet.set_column('P:P', 25, format)
worksheet.set_column('Q:Q', 25, format)
worksheet.set_column('R:R', 25, format)
worksheet.set_column('S:S', 25, format)
worksheet.set_column('T:T', 25, format)
worksheet.set_column('U:U', 25, format)
worksheet.set_column('V:V', 25, format)
worksheet.set_column('W:W', 25, format)
worksheet.set_column('X:X', 25, format)
worksheet.set_column('Y:Y', 25, format)
worksheet.set_column('Z:Z', 25, format)
worksheet.set_column('AP:AP', 25, format)

print('Done!')
# Close the Pandas Excel writer and output the Excel file.
writer.save()

# Open file or folder on OS
path_url = r"Z:\\MIS\\Fone Wasin\\Python\\ULITE SVC TH" 
path_file = path_url + "\*Debt*.xls*"
for filex in glob.iglob(path_file, recursive=True):
    os.path.realpath(path_url)
    # FBI OPEN UP!!!!
    #os.startfile(path_url)
    os.startfile(filex)
    print('Opened File&Folder:', filex)