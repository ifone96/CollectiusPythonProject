#BreakDown
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

# server = 'localhost\sqlexpress' # for a named instance
# server = 'myserver,port' # to specify an alternate port
server = 'collectius-th.crm5.dynamics.com,5558'
database = 'orgbf56918d'
username = 'wasin.k@collectius.com'
password = 'Office365%'
aut = 'ActiveDirectoryPassword'
connect_database = pyodbc.connect(
    'DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+password+';Authentication='+aut)

# DS
print(f"RUN PYTHON file: Breakdown All Account Report.py \n server = collectius-mis.database.windows.net \n database = reporting \n SQL query DS..." )
df_sql_DS = """
--DP
SELECT TOP (100)
	a.alternis_portfolioidname as 'Portfolio',
    a.alternis_batchidname as 'Batch',
    a.alternis_number as 'Account Number',
    a.alternis_invoicenumber as 'Invoice Number',
    a.alternis_accountid as 'UUID',
    a.alternis_contactidname as 'Debtor Name',
    a.alternis_idnumber as 'ID Card',
    a.alternis_processstagename as 'Process Stage',
    a.alternis_outstandingprincipal as 'Outstanding Principal',
    a.alternis_lastpaymentdate as 'Last Payment Date',
    a.alternis_outstandingbalance as 'Outstanding Balance'
FROM alternis_account a
WHERE a.alternis_portfolioidname IN ('AEON1 TH','AEON2 TH','AEON3 TH','KKP1 TH','BMW1 TH','TMB1 TH','TMB2 TH','SME1 TH','SME2 TH','GRAB R2 1 TH')
ORDER BY a.alternis_portfolioidname DESC
            """
df_sql_DS = pd.read_sql(df_sql_DS, connect_database)
print("SQL query DS...is DONE!")


# Set name file with date/times
todaysdate_filename = str(
    datetime.datetime.now().strftime("AllAccountsBreakDownPowerBI'%Y%m%d")) + '.xlsx'
writer = pd.ExcelWriter(todaysdate_filename)
print("Writing File : " + todaysdate_filename)
print("Writing Sheet...")
df_sql_DS.to_excel(writer, index=False, engine='xlsxwriter', sheet_name='DS')
print("...DS")


print("Setting Format...")
workbook = writer.book
worksheet = writer.sheets['DS']
header_format = workbook.add_format({'bold': True})

worksheet.set_row(0, None, header_format)

for column in df_sql_DS:
    column_width = max(df_sql_DS[column].astype(str).map(len).max(), len(column))
    col_idx = df_sql_DS.columns.get_loc(column)
    worksheet.set_column(col_idx, col_idx, column_width)

print("Setting Format...is DONE!")

writer.save()
print("Saved " + todaysdate_filename)
print(todaysdate_filename + " is DONE!")

# Open file or folder on OS
path_url = r"Z:\\MIS\\Fone Wasin\\Python\\Breakdown\\"
path_file = path_url + "\\*.xlsx"
for filex in glob.iglob(path_file, recursive=True):
    os.path.realpath(path_url)
    # os.startfile(path_url)
    os.startfile(filex)
    print('Opened File&Folder:', filex)
