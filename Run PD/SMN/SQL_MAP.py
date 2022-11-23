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
connect_database = pyodbc.connect(
    'DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD=' + password)
sql_cmd = """
        SELECT DISTINCT
        b.alternis_invoicenumber,
        b.alternis_accountid as uuid,
        REPLACE(a.alternis_number,'*','') AS phone,
        a.alternis_phonetypename,
        a.alternis_contactidname,
        b.alternis_idnumber 
        FROM alternis_phone a
        JOIN alternis_account b
        ON a.alternis_contactid = b.alternis_contactid
        WHERE alternis_portfolioidname IN ('SEAMONEY SPL SVC TH','SEAMONEY BCL SVC TH')
        ORDER BY a.alternis_contactidname
            """
df_sql = pd.read_sql(sql_cmd, connect_database)

## Set name file with date/times
#todaysdate_filename = str(
#    datetime.datetime.now().strftime("Leads - SMN %H%M")) + '.xlsx'
writer = pd.ExcelWriter('SQL_MAP.xlsx')

df_sql.to_excel(writer, index=False, engine='xlsxwriter', sheet_name='SQL_MAP')

workbook = writer.book
worksheet2 = writer.sheets['SQL_MAP']

format1 = workbook.add_format({'num_format': '@'})
format2 = workbook.add_format({'num_format': '0000000000'})
header_format = workbook.add_format({'bold': True})

worksheet2.set_row(0, None, header_format)
worksheet2.set_column('A:A', 25)
worksheet2.set_column('B:B', 40)
worksheet2.set_column('C:C', 16)
worksheet2.set_column('D:D', 20)
worksheet2.set_column('E:E', 30)
worksheet2.set_column('F:F', 25)

writer.save()

# Open file or folder on OS
os.startfile("Z:\\MIS\\Fone Wasin\\Python\\Run PD\\SMN\\SQL_MAP.xlsx", 'edit')
print('Opened File&Folder:', "Z:\\MIS\\Fone Wasin\\Python\\Run PD\\SMN\\SQL_MAP.xlsx")