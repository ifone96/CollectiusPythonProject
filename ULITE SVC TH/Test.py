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

count_row = df_combine['Effective Date'].count()