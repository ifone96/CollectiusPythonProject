from ast import If
import datetime
from datetime import datetime, timedelta
import glob
import os
import shutil
from tkinter import HIDDEN
from unittest import skip
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
import openpyxl

data_file_folder = "C:\\Users\\wasin.k\\OneDrive - COLLECTIUS SYSTEMS PTE. LTD\\Ulite\\Payment\\"
df = []
for file in os.listdir(data_file_folder):
    if file.endswith('.xlsx'):
        print('Loading file {0}...'.format(file))
        df.append(pd.read_excel(os.path.join(data_file_folder, file), sheet_name='payment_reconcile_report', header=11))
    if file.endswith('.xlsb'):
        print('Loading file {0}...'.format(file))
        df.append(pd.read_excel(os.path.join(data_file_folder, file), sheet_name='payment_reconcile_report', header=11))
        
# Len(df)
df_combine = pd.concat(df, axis=0)
reCol = {
    'ผู้ซื้อ': 'Account Number',
    'SO No.': 'Invoice/Card Number',
    'ยอดรับชำระรวม(บาท)' : 'Amount',
    'วันที่ชำระ': 'Effective Date',
    'งวด': 'Description'
}

# call rename () method
df_combine.rename(columns=reCol, inplace=True)
df_combine = df_combine[['Account Number','Invoice/Card Number','Amount','Effective Date','Description']]
df_combine = df_combine.assign(**{ 
                                'Account Number+': '',	
                                'Card Number+':'',	
                                'Description+':'',	 
                                'Amount+':'', 	 
                                'Amount Amount in LCY+':'', 
                                'Effective Transaction Date+': '',
                                'Transaction Date Posting+': '=TODAY()',
                                'Payment Channel+': '',
                                'Product Type+': '',
                                'Statement Reference+': ''
                                })
today = str(datetime.today().strftime('%d-%m-%Y'))  
yesterday = datetime.today() - timedelta(days=1)
df_combine = df_combine[df_combine['Effective Date'] == yesterday.strftime('%d-%m-%Y')]
df_combine.dropna(inplace=True)

todaysdate_filename = str(
    datetime.today().strftime("CombineUlitePayment'%Y%m%d")) + '.xlsx'
writer = pd.ExcelWriter(todaysdate_filename)

df_combine.to_excel(writer, index=False, sheet_name= 'Combine_Ulite')
# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['Combine_Ulite']


# Add some cell formats.
format = workbook.add_format({'num_format': '###############'})
format1 = workbook.add_format({'num_format': '@'})
format2 = workbook.add_format({'num_format': '0.00'})
format3 = workbook.add_format({'num_format': 'mm/dd/yyyy'})

# Set the column width and format.
worksheet.set_column('A:A', 30)
worksheet.set_column('B:B', 20)
worksheet.set_column('C:C', 15, format2)
worksheet.set_column('D:D', 20)
worksheet.set_column('E:E', 15)
worksheet.set_column('F:F', 30)
worksheet.set_column('G:G', 30)
worksheet.set_column('H:H', 15)
worksheet.set_column('I:I', 28, format2)
worksheet.set_column('J:J', 28, format2)
worksheet.set_column('K:K', 28, format3)
worksheet.set_column('L:L', 28, format3)
worksheet.set_column('M:M', 25)
worksheet.set_column('N:N', 25)
worksheet.set_column('O:O', 25)
worksheet.set_column('P:P', 25)
worksheet.set_column('Q:Q', 25)

count_row = df_combine['Effective Date'].count()+1
#Formula 
worksheet.write_dynamic_array_formula('F2', '=A2:A'+ str(count_row) +'&""')
worksheet.write_dynamic_array_formula('G2', '=B2:B'+ str(count_row) +'&""')
worksheet.write_dynamic_array_formula('H2', '="No. Install "&E2:E'+ str(count_row) +'')
worksheet.write_dynamic_array_formula('I2', '=C2:C'+ str(count_row) +'*1')
worksheet.write_dynamic_array_formula('J2', '=C2:C'+ str(count_row) +'*1')
worksheet.write_dynamic_array_formula('K2', '=_xlfn.DATE(_xlfn.RIGHT(D2:D'+ str(count_row) +',4),_xlfn.MID(D2:D'+ str(count_row) +',4,2),_xlfn.LEFT(D2:D'+ str(count_row) +',2))')

# Close the Pandas Excel writer and output the Excel file.
import jpype
import asposecells
import shift15m

output_directory = "Examples/SampleFiles/OutputDirectory/"

    # Instantiating a Workbook object

    # Get the first worksheet
worksheet = workbook.getWorksheets().get(0)

# Set values
worksheet.getCells().get(0, 2).setValue(1)
worksheet.getCells().get(1, 2).setValue(2)
worksheet.getCells().get(2, 2).setValue(3)
worksheet.getCells().get(2, 3).setValue(4)
worksheet.getCells().createRange(0, 2, 3, 1).setName("NamedRange")
# Cut and paste cells
cut = worksheet.getCells().createRange("C:C")
worksheet.getCells().insertCutCells(cut, 0, 1, ShiftType.RIGHT)
# Save the excel file.
workbook.save(output_directory + "CutAndPasteCells.xlsx")
writer.save()



src_folder = r"Z:\\MIS\\Fone Wasin\\Python\\ULITE SVC TH\\Payment\\"
dst_folder = r"Z:\\MIS\\Fone Wasin\\Python\\ULITE SVC TH\\Payment\\Uploaded\\"
# move file whose name end with string 'xls'
pattern = src_folder + "*payment*.xls*"
for files in glob.iglob(pattern, recursive=True):
    # extract file name form file path
    file_name = os.path.basename(files)
    #todayy = str(datetime.now().strftime("(Uploaded) %H%M "))
    shutil.move(files, dst_folder + file_name)
    print('Moved:', files)


# Open file or folder on OS
path_url = "Z:\\MIS\\Fone Wasin\\Python\\ULITE SVC TH\\"
path_file = path_url + "\*.xls*"
for filex in glob.iglob(path_file, recursive=True):
    os.path.realpath(path_url)
    # FBI OPEN UP!!!!
    #os.startfile(path_url)
    os.startfile(filex)
    print('Opened File&Folder:', filex)