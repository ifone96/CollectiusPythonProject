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

data_file_folder = 'Z:\\MIS\\Fone Wasin\\Python\\Payment Plan Daily\\OA\\'
sheet = str(datetime.now().strftime('%m-%Y'))
df = []
for file in os.listdir(data_file_folder):
    if file.endswith('.xlsx'):
        print('Loading file {0}...'.format(file))
        df.append(pd.read_excel(os.path.join(data_file_folder, file), sheet_name=sheet))
    if file.endswith('.xlsb'):
        print('Loading file {0}...'.format(file))
        df.append(pd.read_excel(os.path.join(data_file_folder, file), sheet_name=sheet))

len(df)
df_combine = pd.concat(df, axis=0)
reCol = {
    'Portfolio name':       'Portfolio',
    'Date' :                'CreatedDate',
    'Account':              'Account Number',
    'จำนวนงวดที่ชำระ <=36':   'No of Installments',
    'Payment amount':       'Installment Amount',
    'Outstanding':          'Plan Balance',
    'Payment Type':         'Plan Type',
    'ECA':                  'ECA Owner',
    'Name':                 'Mediator Owner'
    
}

# call rename () method
df_combine.rename(columns=reCol, inplace=True)
df_combine = df_combine[[
                        'CreatedDate',
                        'Portfolio',
                        'Account Number',
                        'Mediator Owner',
                        'No of Installments',
                        'Installment Amount',
                        'Plan Balance',
                        'Plan Type',
                        'ECA Owner',
                        ]]

# df_combine.dropna(inplace=True)
today = str(datetime.today().strftime('%Y-%m-%d'))
yesterday = datetime.today() - timedelta(days=1)
#df_combine = df_combine[df_combine['CreatedDate'] == today]

todaysdate_filename = str(
    datetime.today().strftime('PTP_OA_%Y%m%d')) + '.xlsx'
writer = pd.ExcelWriter(todaysdate_filename)

df_combine.to_excel(writer, index=False, sheet_name= 'PTP '+ sheet)
# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['PTP '+ sheet]

# Add some cell formats.
format = workbook.add_format({'num_format': '###############'})
format1 = workbook.add_format({'num_format': '@'})
format2 = workbook.add_format({'num_format': '0.00'})
format3 = workbook.add_format({'num_format': 'dd/mm/yyyy'})


# Set the column width and format.
worksheet.set_column('A:A', 25)
worksheet.set_column('B:B', 25)
worksheet.set_column('C:C', 25)
worksheet.set_column('D:D', 25)
worksheet.set_column('E:E', 25)
worksheet.set_column('F:F', 25, format2)
worksheet.set_column('G:G', 25, format2)
worksheet.set_column('H:H', 25)
worksheet.set_column('I:I', 25)
worksheet.set_column('J:J', 25)
worksheet.set_column('K:K', 25)
worksheet.set_column('L:L', 25)
worksheet.set_column('M:M', 25)
worksheet.set_column('N:N', 25)
worksheet.set_column('O:O', 25)
worksheet.set_column('P:P', 25)
worksheet.set_column('Q:Q', 25)

# Close the Pandas Excel writer and output the Excel file.
writer.save()

# Move file on os base name and path
src_folder = r"Z:\\MIS\\Fone Wasin\\Python\\Payment Plan Daily\\OA\\"
dst_folder = r"Z:\\MIS\\Fone Wasin\\Python\\Payment Plan Daily\\OA\\Uploaded\\"
# move file whose name end with string 'xls'
pattern = src_folder + "\*Daily*.xls*"
for files in glob.iglob(pattern, recursive=True):
    # extract file name form file path
    file_name = os.path.basename(files)
    #todayy = str(datetime.datetime.now().strftime("(Uploaded) %H%M "))
    shutil.move(files, dst_folder + file_name)
    print('Moved:', files)
    break

# Open file or folder on OS
path_url = 'Z:\\MIS\\Fone Wasin\\Python\\Payment Plan Daily\\'
path_file = path_url + '\*OA*.xls*'
for filex in glob.iglob(path_file, recursive=True):
    os.path.realpath(path_url)
    # FBI OPEN UP!!!!
    #os.startfile(path_url)
    os.startfile(filex)
    print('Opened File&Folder:', filex)
    break

path_url = "C:\\Users\\wasin.k\\OneDrive - COLLECTIUS SYSTEMS PTE. LTD\\Desktop"
path_file = path_url + "\*Command MS*.xls*"
for filex in glob.iglob(path_file, recursive=True):
    os.path.realpath(path_url)
    # FBI OPEN UP!!!!
    # os.startfile(path_url)
    os.startfile(filex)
    print('Opened File&Folder:', filex)
    break

source = r"Z:\\MIS\\Fone Wasin\\Python\\Payment Plan Daily\\"
target = "Z:\\MIS\\Report\\Daily\\ECA.xlsx"
s_n = source + "*OA*.xls*"
for files in glob.iglob(s_n, recursive=True):
    file_name = os.path.basename(files)
    shutil.copy(files, target)
    print('Copy&Past:', files)
    break

# Open file or folder on OS
path_url = 'Z:\\MIS\\Report\\Daily\\'
path_file = path_url + '*ECA*.xls*'
for filex in glob.iglob(path_file, recursive=True):
    os.path.realpath(path_url)
    # FBI OPEN UP!!!!
    #os.startfile(path_url)
    os.startfile(filex)
    print('Opened File&Folder:', filex)
    break
