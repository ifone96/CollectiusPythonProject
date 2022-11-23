from ast import If
import datetime
from datetime import datetime, timedelta
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
import numpy as np
import pyodbc
import xlsxwriter
import axis
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

# Mai
print(f"SQL query All Stage...")
sql_Mai = """
            --All Stage
            SELECT a.alternis_portfolioidname as 'Portfolio',
                a.alternis_number as 'Account Number',
                a.alternis_invoicenumber as 'Invoice Number',
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
df_sql_Mai = pd.read_sql(sql_Mai, connect_database)
print("SQL query DP All Stage...is DONE!")

# DP
print(f"SQL query DP...")
sql_DP = """
            --DP
            SELECT 
                a.alternis_portfolioidname as 'Portfolio',
                a.alternis_number as 'Account Number',
                a.alternis_invoicenumber as 'Invoice Number',
                a.alternis_accountid as 'uuid',
                phone.alternis_number as 'Phone Number',
                phone.alternis_phonetypename as 'Phone Type',
                a.alternis_contactidname as 'Debtor Name',
                a.alternis_idnumber as 'ID Card',
                phone.alternis_verificationstatusname as 'Verification Status',
                a.alternis_processstagename as 'Process Stage',
                a.alternis_outstandingprincipal as 'Outstanding Principal',
                a.alternis_lastpaymentdate as 'Last Payment Date',
                a.alternis_outstandingbalance as 'Outstanding Balance',
                phone_call.subject as 'PhoneCall Subject',
                phone_call.alternis_contactdispositionname as 'Contact Disposition',
                phone_call.alternis_calloutcomename as 'Calloutcome',
                phone_call.createdon as 'Last Phonecall Createdon',
                datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) as 'Last Touch Day',
            CASE WHEN datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >40000 then 'No Activity'
            WHEN datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >120 then '07. More than 4 Months'
            WHEN datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >90 then '06. 3 Monhts to 4 Months'
            WHEN datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >60 then '05. 2 Months to 3 Months'
            WHEN datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >30 then '04. 1 Month to 2 Months'
            WHEN datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >21 then '03. 3 Weeks to 1 Months'
            WHEN datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >14 then '02. 2 Weeks to 3 Weeks'
            WHEN datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >7 then '01. 1 Week to 2 Weeks'
            WHEN datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >=0 then '00. Less Than 1 Week'
            else 'No Activity'
            end as Last_Touch
            FROM alternis_account a
                LEFT JOIN alternis_phone phone on phone.alternis_contactid = a.alternis_contactid
                LEFT JOIN phonecall phone_call on phone_call.phonenumber = phone.alternis_number and phone_call.regardingobjectid = a.alternis_accountid and phone_call.activityid = (SELECT TOP(1) activityid
                    FROM [phonecall] phoneCall
                    WHERE phoneCall.phonenumber = phone.alternis_number and phoneCall.regardingobjectid = a.alternis_accountid
                    ORDER BY phoneCall.createdon DESC)
            WHERE a.alternis_portfolioidname IN ('KKP1 TH','BMW1 TH','TMB1 TH','TMB2 TH','SME1 TH','SME2 TH','GRAB R2 1 TH')
            AND a.alternis_processstagename NOT IN ('Closed','Pending Close Review','Pending Paid Review')
            ORDER BY a.alternis_portfolioidname, phone_call.createdon DESC
            """
df_sql_DP = pd.read_sql(sql_DP, connect_database)
print("SQL query DP...is DONE!")

serverDWH = 'collectiusdwhph.database.windows.net'
databaseDWH = 'dwh_th_2022'
usernameDWH = 'atiwat'
passwordDWH = '2a#$dfERat^%'
connect_databaseDWH = pyodbc.connect(
    'DRIVER={ODBC Driver 17 for SQL Server};SERVER='+serverDWH+';DATABASE='+databaseDWH+';UID='+usernameDWH+';PWD=' + passwordDWH)

# AEON
print(f"SQL query AEON...")
sql_AEON = """
                --AEON on Local dwh_th_2022
                select a.[alternis_portfolioidname] as 'Portfolio',
                    a.alternis_number as 'Account Number',
                    cast(a.alternis_invoicenumber as text) as 'Invoice Number',
                    a.[alternis_accountid] as uuid,
                    phone.alternis_number as 'Phone Number',
                    phone.alternis_phonetypename as 'Phone Type',
                    a.[alternis_contactidname] as 'Debtor Name',
                    a.alternis_idnumber as 'ID Card',
                    phone.alternis_verificationstatusname as 'Verification Status',
                    a.[alternis_processstagename] as 'Process Stage',
                    a.alternis_outstandingprincipal as 'Outstanding Principal',
                    a.alternis_lastpaymentdate as 'Last Payment Date',
                    a.alternis_outstandingbalance as 'Outstanding Balance',
                    phone_call.subject as 'PhoneCall Subject',
                    phone_call.alternis_contactdispositionname as 'Contact Disposition',
                    phone_call.alternis_calloutcomename as 'Calloutcome',
                    phone_call.createdon as 'Last Phonecall Createdon',
                    datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) as 'Last Touch Day',
                    case when datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >40000 then 'No Activity'
                	when datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >120 then '07. More than 4 Months'
                	when datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >90 then '06. 3 Monhts to 4 Months'
                	when datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >60 then '05. 2 Months to 3 Months'
                	when datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >30 then '04. 1 Month to 2 Months'
                	when datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >21 then '03. 3 Weeks to 1 Months'
                	when datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >14 then '02. 2 Weeks to 3 Weeks'
                	when datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >7 then '01. 1 Week to 2 Weeks'
                	when datediff(day,floor(cast(phone_call.createdon as float)),CAST( GETDATE() AS Date )) >=0 then '00. Less Than 1 Week' else 'No Activity' end as Last_Touch
                FROM stage.alternis_account a
                    join stage.contact c on c.contactid = a.alternis_contactid
                    left join stage.alternis_phone phone on phone.alternis_contactid = a.alternis_contactid
                    left join stage.phonecall phone_call on phone_call.phonenumber = phone.alternis_number and phone_call.regardingobjectid = a.alternis_accountid and phone_call.activityid = (SELECT TOP(1)
                            activityid
                        FROM [stage].[phonecall] phoneCall
                        where phoneCall.phonenumber = phone.alternis_number and phoneCall.regardingobjectid = a.alternis_accountid
                        ORDER BY phoneCall.createdon DESC)
                    left join stage.task on task.regardingobjectid = a.alternis_accountid and task.activityid = (select top(1)
                            activityid
                        from stage.task tas
                        where tas.regardingobjectid = a.alternis_accountid
                        ORDER BY tas.createdon DESC)
                --WHERE a.alternis_portfolioidname IN ('AEON1 TH','AEON2 TH','AEON3 TH','KKP1 TH','BMW1 TH','TMB1 TH','TMB2 TH','SME1 TH','SME2 TH') AND a.alternis_processstagename NOT IN ('Closed')
                where a.alternis_portfolioidname IN ('AEON1 TH','AEON2 TH','AEON3 TH') AND a.alternis_processstagename NOT IN ('Closed','Pending Close Review','Pending Paid Review')
                ORDER BY a.alternis_portfolioidname, phone_call.createdon DESC
            """
df_sql_AEON = pd.read_sql(sql_AEON, connect_databaseDWH)
print("SQL query AEON...is DONE!")

# Set name file with date/times
todaysdate_filename = str(
    datetime.now().strftime("All_DP_Dynamics'DHW_%Y%m%d")) + '.xlsx'
writer = pd.ExcelWriter(todaysdate_filename)
print("Writing File : " + todaysdate_filename)

print("Writing Sheet...")
df_sql_DP.to_excel(writer, index=False, engine='xlsxwriter', sheet_name='DP')
df_sql_AEON.to_excel(writer, index=False, engine='xlsxwriter', sheet_name='AEON')
df_sql_Mai.to_excel(writer, index=False, engine='xlsxwriter', sheet_name='All Stage IH')
print("...All Stage")


print("Setting Format...")
workbook = writer.book
worksheet = writer.sheets['All Stage IH']
worksheet2 = writer.sheets['DP']
worksheet3 = writer.sheets['AEON']
header_format = workbook.add_format({'bold': True})

worksheet.set_row(0, None, header_format)
for column in df_sql_Mai:
    column_width = max(df_sql_Mai[column].astype(str).map(len).max(), len(column))
    col_idx = df_sql_Mai.columns.get_loc(column)
    worksheet.set_column(col_idx, col_idx, column_width)
    
worksheet2.set_row(0, None, header_format)  
for column in df_sql_DP:
    column_width = max(df_sql_DP[column].astype(str).map(len).max(), len(column))
    col_idx = df_sql_DP.columns.get_loc(column)
    worksheet2.set_column(col_idx, col_idx, column_width)\

worksheet3.set_row(0, None, header_format)  
for column in df_sql_AEON:
    column_width = max(df_sql_AEON[column].astype(str).map(len).max(), len(column))
    col_idx = df_sql_AEON.columns.get_loc(column)
    worksheet3.set_column(col_idx, col_idx, column_width)

print("Setting Format...is DONE!")

writer.save()
print("Saved " + todaysdate_filename)
print(todaysdate_filename + " is DONE!")

src_folder = "Z:\\MIS\\Fone Wasin\\Python\\Breakdown\\"
dst_folder = "C:\\Users\\wasin.k\\OneDrive - COLLECTIUS SYSTEMS PTE. LTD\\Documents - MIS-TH\\Morning Data"
pattern = src_folder + "\\*DP*.xls*"
for files in glob.iglob(pattern, recursive=True):
    file_name = os.path.basename(files)
    shutil.copy(files, dst_folder) 
    print('Moved:', files)
    break

src_folder = r"Z:\\MIS\\Fone Wasin\\Python\\Breakdown\\"
dst_folder = r"Z:\\MIS\\Fone Wasin\\Python\\Breakdown\\Uploaded\\" 
# move file whose name end with string 'xls'
pattern = src_folder + "*DP*.xls*"
for files in glob.iglob(pattern, recursive=True):
    # extract file name form file path
    file_name = os.path.basename(files)
    #todayy = str(datetime.now().strftime("(Uploaded) %H%M "))
    shutil.move(files, dst_folder + file_name)
    print('Moved:', files)
    break

path_url = "C:\\Users\\wasin.k\\OneDrive - COLLECTIUS SYSTEMS PTE. LTD\\Documents - MIS-TH\\Morning Data\\"
path_file = path_url + "*DP*.xls*"
for filex in glob.iglob(path_file, recursive=True):
    os.path.realpath(path_url)
    # FBI OPEN UP!!!!
    # os.startfile(path_url)
    os.startfile(filex)
    print('Opened File&Folder:', filex)
    break

