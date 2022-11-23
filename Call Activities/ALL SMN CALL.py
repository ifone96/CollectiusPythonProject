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

###SMN
data_file_folder = "Z:\\MIS\\Fone Wasin\\Python\\Call Activities\\BASESMN\\"
df = []
for file in os.listdir(data_file_folder):
    if file.endswith(".xlsx"):
        print('Loading file {0}...'.format(file))
        df.append(pd.read_excel(os.path.join(
            data_file_folder, file), sheet_name="BaseSMN"))
    if file.endswith(".xlsb"):
        print('Loading file {0}...'.format(file))
        df.append(pd.read_excel(os.path.join(
            data_file_folder, file), sheet_name="BaseSMN"))

df_combine = pd.concat(df, axis=0)
reCol = {
    'user_id': 'user_id',
    'Name': 'user_name',
    'phone_number': 'user_phone',
    'oa_code': 'Code ทีม',
    'user_type': 'user_type',
    'assign_type': 'destination_type',
    'aging': 'aging',
    'total_amount_to_pay': 'total_outstanding',
    'รอบบิล': 'due_date'
        }
# call rename () method
df_combine.rename(columns=reCol, inplace=True)

df_combine = df_combine[[
                        'loan_type',
                        'company_name',
                        'Code ทีม',
                        'user_id',
                        'user_name',
                        'user_phone',
                        'user_type',
                        'aging',
                        'due_date',
                        'total_outstanding',
                        'destination_type'
                        ]].astype('string')

df_combine = df_combine.assign(**{
                        'result_code': '', 
                        'sub_reason': '', 
                        'ptp_date': '', 
                        'ptp_amount ': '', 
                        'contact_result': '', 
                        'collection_agent': '', 
                        'outbound_number': '', 
                        'result_date': '', 
                        'result_time': '',
                        'PhoneF(x)': ''
                        })

df_combine.insert(2, 'No', '', True)
df_combine.insert(3, 'report_date', '=Today()-1', True)
df_combine = df_combine[df_combine['loan_type'] == 'SPayLater']
df_combine = df_combine[df_combine['company_name'] == 'SMN']
df_combine.drop_duplicates(subset='user_id', inplace=True)

print(df_combine)

today = datetime.today()
yesterday = datetime.today() - timedelta(days=1)

# server = 'localhost\sqlexpress' # for a named instance
# server = 'myserver,port' # to specify an alternate port
server = 'collectius-th.crm5.dynamics.com,5558'
database = 'orgbf56918d'
username = 'wasin.k@collectius.com'
password = 'Office365%'
aut = 'ActiveDirectoryPassword'
connect_database = pyodbc.connect(
    'DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+password+';Authentication='+aut)
# LPC
print(f"SQL Qury")
sql_cmd_LPC =   yesterday.strftime(
                
                """
                
                --Last Call
                SELECT

                	a.alternis_portfolioidname as 'Portfolio',
                	a.alternis_number as 'Account Number',
                	a.alternis_invoicenumber as 'Invoice',
                	a.alternis_contactidname as 'Name',
                	phone.alternis_phonetypename as 'Phone Type',
                	phone_call.phonenumber as 'Phone Number',
                	phone_call.alternis_calloutcomename as 'Call Outcome',
                	phone_call.alternis_contactdispositionname as 'Contact Disposition',
                	phone_call.description as 'Description',
                	phone_call.createdon as 'Last Phonecall Createdon',
                	phone_call.actualdurationminutes as 'Duration',
                	phone_call.subject as 'Subject',
                	phone_call.modifiedbyname as 'Agent Call'

                FROM alternis_account a
                FULL JOIN alternis_phone phone on phone.alternis_contactid = a.alternis_contactid
                FULL JOIN phonecall phone_call on phone_call.phonenumber = phone.alternis_number and phone_call.regardingobjectid = a.alternis_accountid
                WHERE a.alternis_portfolioidname IN ('SEAMONEY SPL SVC TH','SEAMONEY BCL SVC TH')
                AND phone_call.createdon >= '%Y-%m-%d 00:00:00.000'
                ORDER BY phone_call.createdon DESC      
                          
                """
                )

sql_cmd_LPC = pd.read_sql(sql_cmd_LPC, connect_database)
print("SQL query sql_cmd_LPC...is DONE!")

#   PTP
print("SQL query DP...")
sql_cmd_PTP =   yesterday.strftime(
                """

                --PTP
                SELECT

                	a.alternis_portfolioidname as 'Portfolio'
                	,a.alternis_contactidname as 'Name'
                	,a.alternis_number as 'Account Number'
                	,a.alternis_invoicenumber as 'Invoice Number'
                	,p.alternis_firstpaymentdate as '1st Payment Date'
                	,p.alternis_installmentamount as 'Installment Amount'
                    ,p.alternis_amountoninstallments as 'Total Amount on Installment'			
                    ,p.alternis_totaldiscountvalue	as 'Total Discount Value'
                	,p.alternis_amountpaid 'Paid'
                	,p.statuscode as 'Status Reason'
                	,p.createdon as 'Created On'
                	,p.alternis_paymentplanid 'PTP ID'

                FROM alternis_account a 
                INNER JOIN alternis_paymentplan p ON p.alternis_accountid = a.alternis_accountid
                WHERE a.alternis_portfolioidname IN ('SEAMONEY SPL SVC TH','SEAMONEY BCL SVC TH')
                AND p.createdon >= '%Y-%m-%d 00:00:00.000'
                ORDER BY p.createdon DESC
                
                """
                )

sql_cmd_PTP = pd.read_sql(sql_cmd_PTP, connect_database)
print("SQL query sql_cmd_PTP...is DONE!")

# Set name file with date/times
todaysdate_filename = yesterday.strftime(("SPayLater-SMN Daily report as of %Y-%m-%d")) + '.xlsx'
writer = pd.ExcelWriter(todaysdate_filename)
print("Writing File : " + todaysdate_filename)


print("Writing Sheet...")
df_combine.to_excel(writer, index=False, 
                    engine='xlsxwriter', sheet_name='SPayLater-SMN')
print("...SMN")
sql_cmd_LPC.to_excel(writer, index=False,
                     engine='xlsxwriter', sheet_name='LPC')
print("...LPC")
sql_cmd_PTP.to_excel(writer, index=False,
                     engine='xlsxwriter', sheet_name='PTP')
print("...PTP")


print("Setting Format...")
workbook = writer.book
worksheet = writer.sheets['LPC']
worksheet2 = writer.sheets['PTP']
worksheet3 = writer.sheets['SPayLater-SMN']
header_format = workbook.add_format({'bold': True})

worksheet.set_row(0, None, header_format)
worksheet2.set_row(0, None, header_format)
worksheet3.set_row(0, None, header_format)

# Add some cell formats.
format = workbook.add_format({'num_format': '0000000000'})
format1 = workbook.add_format({'num_format': '@'})
format2 = workbook.add_format({'num_format': '0.00'})
format3 = workbook.add_format({'num_format': 'yyyy-mm-dd'})

# Set the column width and format.
worksheet3.set_column('A:A', 15)
worksheet3.set_column('B:B', 15)
worksheet3.set_column('C:C', 10)
worksheet3.set_column('D:D', 25, format3)
worksheet3.set_column('E:E', 25)
worksheet3.set_column('F:F', 25)
worksheet3.set_column('G:G', 25)
worksheet3.set_column('H:H', 25, format1)
worksheet3.set_column('I:I', 25)
worksheet3.set_column('J:J', 25)
worksheet3.set_column('K:K', 25)
worksheet3.set_column('L:L', 25)
worksheet3.set_column('M:M', 25)
worksheet3.set_column('N:N', 25)
worksheet3.set_column('O:O', 25)
worksheet3.set_column('P:P', 25, format3)
worksheet3.set_column('Q:Q', 25, format2)
worksheet3.set_column('R:R', 25)
worksheet3.set_column('S:S', 25)
worksheet3.set_column('T:T', 25)
worksheet3.set_column('U:U', 25, format3)
worksheet3.set_column('V:V', 25)
worksheet3.set_column('W:W', 15)

#Formula 
worksheet3.write_dynamic_array_formula('W2','=REPT(0,1)&(RIGHT(H2,9))')
worksheet3.write_dynamic_array_formula('N2','=_xlfn.XLOOKUP($F2,LPC!$B:$B,LPC!$G:$G)')
worksheet3.write_dynamic_array_formula('O2','=_xlfn.XLOOKUP($F2,LPC!$B:$B,LPC!$H:$H)')
worksheet3.write_dynamic_array_formula('P2','=_xlfn.XLOOKUP($F2,PTP!$C:$C,PTP!$E:$E)')
worksheet3.write_dynamic_array_formula('Q2','=_xlfn.XLOOKUP($F2,PTP!$C:$C,PTP!$F:$F)')
worksheet3.write_dynamic_array_formula('R2','=_xlfn.XLOOKUP($F2,LPC!$B:$B,LPC!$L:$L)')
worksheet3.write_dynamic_array_formula('S2','=_xlfn.XLOOKUP($F2,LPC!$B:$B,LPC!$M:$M)')
worksheet3.write_dynamic_array_formula('T2',"=_xlfn.INDEX('Z:\\MIS\\Fone Wasin\\Python\\Call Activities\\[Seamoney Outbond Phone.xlsx]Sheet1'!$A:$A,RANDBETWEEN(1,COUNTA('Z:\\MIS\Fone Wasin\\Python\\Call Activities\\[Seamoney Outbond Phone.xlsx]Sheet1'!$A:$A)),1)")
worksheet3.write_dynamic_array_formula('U2','=Today()-1')
worksheet3.write_dynamic_array_formula('V2','=_xlfn.XLOOKUP($F2,LPC!$B:$B,LPC!$K:$K)')

print("Setting Format...is DONE!")
print("Saved " + todaysdate_filename)
print("Waitong For Open File: " + todaysdate_filename)
writer.save()

print(todaysdate_filename + " is DONE!")
# Open file or folder on OS
path_url = "Z:\\MIS\\Fone Wasin\\Python\\Call Activities\\"
path_file = path_url + "\*SPayLater-SMN Daily report as of*.xls*"
for filex in glob.iglob(path_file, recursive=True):
    os.path.realpath(path_url)
    # FBI OPEN UP!!!!
    # os.startfile(path_url)
    os.startfile(filex)
    print('Opened File&Folder:', filex)



###UNC
data_file_folder = "Z:\\MIS\\Fone Wasin\\Python\\Call Activities\\BASESMN\\"
df = []
for file in os.listdir(data_file_folder):
    if file.endswith(".xlsx"):
        print('Loading file {0}...'.format(file))
        df.append(pd.read_excel(os.path.join(
            data_file_folder, file), sheet_name="BaseSMN"))
    if file.endswith(".xlsb"):
        print('Loading file {0}...'.format(file))
        df.append(pd.read_excel(os.path.join(
            data_file_folder, file), sheet_name="BaseSMN"))

df_combine = pd.concat(df, axis=0)
reCol = {
    'user_id': 'user_id',
    'Name': 'user_name',
    'phone_number': 'user_phone',
    'oa_code': 'Code ทีม',
    'user_type': 'user_type',
    'assign_type': 'destination_type',
    'aging': 'aging',
    'total_amount_to_pay': 'total_outstanding',
    'รอบบิล': 'due_date'
        }
# call rename () method
df_combine.rename(columns=reCol, inplace=True)

df_combine = df_combine[[
                        'loan_type',
                        'company_name',
                        'Code ทีม',
                        'user_id',
                        'user_name',
                        'user_phone',
                        'user_type',
                        'aging',
                        'due_date',
                        'total_outstanding',
                        'destination_type'
                        ]].astype('string')

df_combine = df_combine.assign(**{
                        'result_code': '', 
                        'sub_reason': '', 
                        'ptp_date': '', 
                        'ptp_amount ': '', 
                        'contact_result': '', 
                        'collection_agent': '', 
                        'outbound_number': '', 
                        'result_date': '', 
                        'result_time': '',
                        'PhoneF(x)': ''
                        })

df_combine.insert(2, 'No', '', True)
df_combine.insert(3, 'report_date', '=Today()-1', True)
df_combine = df_combine[df_combine['loan_type'] == 'SPayLater']
df_combine = df_combine[df_combine['company_name'] == 'UNC']
df_combine.drop_duplicates(subset='user_id', inplace=True)
print(df_combine)

# LPC
print(f"RUN PYTHON file:")
sql_cmd_LPC =   yesterday.strftime(
                
                """
                
                --Last Call
                SELECT

                	a.alternis_portfolioidname as 'Portfolio',
                	a.alternis_number as 'Account Number',
                	a.alternis_invoicenumber as 'Invoice',
                	a.alternis_contactidname as 'Name',
                	phone.alternis_phonetypename as 'Phone Type',
                	phone_call.phonenumber as 'Phone Number',
                	phone_call.alternis_calloutcomename as 'Call Outcome',
                	phone_call.alternis_contactdispositionname as 'Contact Disposition',
                	phone_call.description as 'Description',
                	phone_call.createdon as 'Last Phonecall Createdon',
                	phone_call.actualdurationminutes as 'Duration',
                	phone_call.subject as 'Subject',
                	phone_call.modifiedbyname as 'Agent Call'

                FROM alternis_account a
                FULL JOIN alternis_phone phone on phone.alternis_contactid = a.alternis_contactid
                FULL JOIN phonecall phone_call on phone_call.phonenumber = phone.alternis_number and phone_call.regardingobjectid = a.alternis_accountid
                WHERE a.alternis_portfolioidname IN ('SEAMONEY SPL SVC TH','SEAMONEY BCL SVC TH')
                AND phone_call.createdon >= '%Y-%m-%d 00:00:00.000'
                ORDER BY phone_call.createdon DESC      
                          
                """
                )

sql_cmd_LPC = pd.read_sql(sql_cmd_LPC, connect_database)
print("SQL query sql_cmd_LPC...is DONE!")

#   PTP
print("SQL query DP...")
sql_cmd_PTP =   yesterday.strftime(
                """

                --PTP
                SELECT

                	a.alternis_portfolioidname as 'Portfolio'
                	,a.alternis_contactidname as 'Name'
                	,a.alternis_number as 'Account Number'
                	,a.alternis_invoicenumber as 'Invoice Number'
                	,p.alternis_firstpaymentdate as '1st Payment Date'
                	,p.alternis_installmentamount as 'Installment Amount'
                    ,p.alternis_amountoninstallments as 'Total Amount on Installment'			
                    ,p.alternis_totaldiscountvalue	as 'Total Discount Value'
                	,p.alternis_amountpaid 'Paid'
                	,p.statuscode as 'Status Reason'
                	,p.createdon as 'Created On'
                	,p.alternis_paymentplanid 'PTP ID'

                FROM alternis_account a 
                INNER JOIN alternis_paymentplan p ON p.alternis_accountid = a.alternis_accountid
                WHERE a.alternis_portfolioidname IN ('SEAMONEY SPL SVC TH','SEAMONEY BCL SVC TH')
                AND p.createdon >= '%Y-%m-%d 00:00:00.000'
                ORDER BY p.createdon DESC
                
                """
                )

sql_cmd_PTP = pd.read_sql(sql_cmd_PTP, connect_database)
print("SQL query sql_cmd_PTP...is DONE!")


# Set name file with date/times
todaysdate_filename = yesterday.strftime("SPayLater-UNC Daily report as of %Y-%m-%d") + '.xlsx'
writer = pd.ExcelWriter(todaysdate_filename)
print("Writing File : " + todaysdate_filename)


print("Writing Sheet...")
df_combine.to_excel(writer, index=False, 
                    engine='xlsxwriter', sheet_name='SPayLater-UNC')
print("...SMN")
sql_cmd_LPC.to_excel(writer, index=False,
                     engine='xlsxwriter', sheet_name='LPC')
print("...LPC")
sql_cmd_PTP.to_excel(writer, index=False,
                     engine='xlsxwriter', sheet_name='PTP')
print("...PTP")


print("Setting Format...")
workbook = writer.book
worksheet = writer.sheets['LPC']
worksheet2 = writer.sheets['PTP']
worksheet3 = writer.sheets['SPayLater-UNC']
header_format = workbook.add_format({'bold': True})

worksheet.set_row(0, None, header_format)
worksheet2.set_row(0, None, header_format)
worksheet3.set_row(0, None, header_format)

# Add some cell formats.
format = workbook.add_format({'num_format': '0000000000'})
format1 = workbook.add_format({'num_format': '@'})
format2 = workbook.add_format({'num_format': '0.00'})
format3 = workbook.add_format({'num_format': 'yyyy-mm-dd'})

# Set the column width and format.
worksheet3.set_column('A:A', 15)
worksheet3.set_column('B:B', 15)
worksheet3.set_column('C:C', 10)
worksheet3.set_column('D:D', 25, format3)
worksheet3.set_column('E:E', 25)
worksheet3.set_column('F:F', 25)
worksheet3.set_column('G:G', 25)
worksheet3.set_column('H:H', 25, format1)
worksheet3.set_column('I:I', 25)
worksheet3.set_column('J:J', 25)
worksheet3.set_column('K:K', 25)
worksheet3.set_column('L:L', 25)
worksheet3.set_column('M:M', 25)
worksheet3.set_column('N:N', 25)
worksheet3.set_column('O:O', 25)
worksheet3.set_column('P:P', 25, format3)
worksheet3.set_column('Q:Q', 25, format2)
worksheet3.set_column('R:R', 25)
worksheet3.set_column('S:S', 25)
worksheet3.set_column('T:T', 25)
worksheet3.set_column('U:U', 25, format3)
worksheet3.set_column('V:V', 25)
worksheet3.set_column('W:W', 15)

#Formula 
worksheet3.write_dynamic_array_formula('W2','=REPT(0,1)&(RIGHT(H2,9))')
worksheet3.write_dynamic_array_formula('N2','=_xlfn.XLOOKUP($F2,LPC!$B:$B,LPC!$G:$G)')
worksheet3.write_dynamic_array_formula('O2','=_xlfn.XLOOKUP($F2,LPC!$B:$B,LPC!$H:$H)')
worksheet3.write_dynamic_array_formula('P2','=_xlfn.XLOOKUP($F2,PTP!$C:$C,PTP!$E:$E)')
worksheet3.write_dynamic_array_formula('Q2','=_xlfn.XLOOKUP($F2,PTP!$C:$C,PTP!$F:$F)')
worksheet3.write_dynamic_array_formula('R2','=_xlfn.XLOOKUP($F2,LPC!$B:$B,LPC!$L:$L)')
worksheet3.write_dynamic_array_formula('S2','=_xlfn.XLOOKUP($F2,LPC!$B:$B,LPC!$M:$M)')
worksheet3.write_dynamic_array_formula('T2',"=_xlfn.INDEX('Z:\\MIS\\Fone Wasin\\Python\\Call Activities\\[Seamoney Outbond Phone.xlsx]Sheet1'!$A:$A,RANDBETWEEN(1,COUNTA('Z:\\MIS\Fone Wasin\\Python\\Call Activities\\[Seamoney Outbond Phone.xlsx]Sheet1'!$A:$A)),1)")
worksheet3.write_dynamic_array_formula('U2','=Today()-1')
worksheet3.write_dynamic_array_formula('V2','=_xlfn.XLOOKUP($F2,LPC!$B:$B,LPC!$K:$K)')

print("Setting Format...is DONE!")
print("Saved " + todaysdate_filename)
print("Waitong For Open File: " + todaysdate_filename)
writer.save()

print(todaysdate_filename + " is DONE!")
# Open file or folder on OS
path_url = "Z:\\MIS\\Fone Wasin\\Python\\Call Activities\\"
path_file = path_url + "\*SPayLater-UNC Daily report as of*.xls*"
for filex in glob.iglob(path_file, recursive=True):
    os.path.realpath(path_url)
    # FBI OPEN UP!!!!
    # os.startfile(path_url)
    os.startfile(filex)
    print('Opened File&Folder:', filex)



###BCL
data_file_folder = "Z:\\MIS\\Fone Wasin\\Python\\Call Activities\\BASESMN\\"
df = []
for file in os.listdir(data_file_folder):
    if file.endswith(".xlsx"):
        print('Loading file {0}...'.format(file))
        df.append(pd.read_excel(os.path.join(
            data_file_folder, file), sheet_name="BaseSMN"))
    if file.endswith(".xlsb"):
        print('Loading file {0}...'.format(file))
        df.append(pd.read_excel(os.path.join(
            data_file_folder, file), sheet_name="BaseSMN"))

df_combine = pd.concat(df, axis=0)
reCol = {
    'user_id': 'user_id',
    'Name': 'user_name',
    'phone_number': 'user_phone',
    'oa_code': 'Code ทีม',
    'user_type': 'user_type',
    'assign_type': 'destination_type',
    'aging': 'aging',
    'total_amount_to_pay': 'total_outstanding',
    'รอบบิล': 'due_date'
        }
# call rename () method
df_combine.rename(columns=reCol, inplace=True)

df_combine = df_combine[[
                        'loan_type',
                        'company_name',
                        'Code ทีม',
                        'user_id',
                        'user_name',
                        'user_phone',
                        'user_type',
                        'aging',
                        'due_date',
                        'total_outstanding',
                        'destination_type'
                        ]].astype('string')

df_combine = df_combine.assign(**{
                        'result_code': '', 
                        'sub_reason': '', 
                        'ptp_date': '', 
                        'ptp_amount ': '', 
                        'contact_result': '', 
                        'collection_agent': '', 
                        'outbound_number': '', 
                        'result_date': '', 
                        'result_time': '',
                        'PhoneF(x)': ''
                        })

df_combine.insert(2, 'No', '', True)
df_combine.insert(3, 'report_date', '=Today()-1', True)
df_combine = df_combine[df_combine['loan_type'] == 'SEasyCash']
df_combine = df_combine[df_combine['company_name'] == 'SMN']
df_combine.drop_duplicates(subset='user_id', inplace=True)
print(df_combine)

# LPC
print(f"RUN PYTHON file:")
sql_cmd_LPC =   yesterday.strftime(
                
                """
                
                --Last Call
                SELECT

                	a.alternis_portfolioidname as 'Portfolio',
                	a.alternis_number as 'Account Number',
                	a.alternis_invoicenumber as 'Invoice',
                	a.alternis_contactidname as 'Name',
                	phone.alternis_phonetypename as 'Phone Type',
                	phone_call.phonenumber as 'Phone Number',
                	phone_call.alternis_calloutcomename as 'Call Outcome',
                	phone_call.alternis_contactdispositionname as 'Contact Disposition',
                	phone_call.description as 'Description',
                	phone_call.createdon as 'Last Phonecall Createdon',
                	phone_call.actualdurationminutes as 'Duration',
                	phone_call.subject as 'Subject',
                	phone_call.modifiedbyname as 'Agent Call'

                FROM alternis_account a
                FULL JOIN alternis_phone phone on phone.alternis_contactid = a.alternis_contactid
                FULL JOIN phonecall phone_call on phone_call.phonenumber = phone.alternis_number and phone_call.regardingobjectid = a.alternis_accountid
                WHERE a.alternis_portfolioidname IN ('SEAMONEY SPL SVC TH','SEAMONEY BCL SVC TH')
                AND phone_call.createdon >= '%Y-%m-%d 00:00:00.000'
                ORDER BY phone_call.createdon DESC      
                          
                """
                )

sql_cmd_LPC = pd.read_sql(sql_cmd_LPC, connect_database)
print("SQL query sql_cmd_LPC...is DONE!")

#   PTP
print("SQL query DP...")
sql_cmd_PTP =   yesterday.strftime(
                """

                --PTP
                SELECT

                	a.alternis_portfolioidname as 'Portfolio'
                	,a.alternis_contactidname as 'Name'
                	,a.alternis_number as 'Account Number'
                	,a.alternis_invoicenumber as 'Invoice Number'
                	,p.alternis_firstpaymentdate as '1st Payment Date'
                	,p.alternis_installmentamount as 'Installment Amount'
                    ,p.alternis_amountoninstallments as 'Total Amount on Installment'			
                    ,p.alternis_totaldiscountvalue	as 'Total Discount Value'
                	,p.alternis_amountpaid 'Paid'
                	,p.statuscode as 'Status Reason'
                	,p.createdon as 'Created On'
                	,p.alternis_paymentplanid 'PTP ID'

                FROM alternis_account a 
                INNER JOIN alternis_paymentplan p ON p.alternis_accountid = a.alternis_accountid
                WHERE a.alternis_portfolioidname IN ('SEAMONEY SPL SVC TH','SEAMONEY BCL SVC TH')
                AND p.createdon >= '%Y-%m-%d 00:00:00.000'
                ORDER BY p.createdon DESC
                
                """
                )

sql_cmd_PTP = pd.read_sql(sql_cmd_PTP, connect_database)
print("SQL query sql_cmd_PTP...is DONE!")


# Set name file with date/times
todaysdate_filename = yesterday.strftime(("SEasyCash-BCL Daily report as of %Y-%m-%d")) + '.xlsx'
writer = pd.ExcelWriter(todaysdate_filename)
print("Writing File : " + todaysdate_filename)


print("Writing Sheet...")
df_combine.to_excel(writer, index=False, 
                    engine='xlsxwriter', sheet_name='SEasyCash-BCL')
print("...SMN")
sql_cmd_LPC.to_excel(writer, index=False,
                     engine='xlsxwriter', sheet_name='LPC')
print("...LPC")
sql_cmd_PTP.to_excel(writer, index=False,
                     engine='xlsxwriter', sheet_name='PTP')
print("...PTP")


print("Setting Format...")
workbook = writer.book
worksheet = writer.sheets['LPC']
worksheet2 = writer.sheets['PTP']
worksheet3 = writer.sheets['SEasyCash-BCL']
header_format = workbook.add_format({'bold': True})

worksheet.set_row(0, None, header_format)
worksheet2.set_row(0, None, header_format)
worksheet3.set_row(0, None, header_format)

# Add some cell formats.
format = workbook.add_format({'num_format': '0000000000'})
format1 = workbook.add_format({'num_format': '@'})
format2 = workbook.add_format({'num_format': '0.00'})
format3 = workbook.add_format({'num_format': 'yyyy-mm-dd'})

# Set the column width and format.
worksheet3.set_column('A:A', 15)
worksheet3.set_column('B:B', 15)
worksheet3.set_column('C:C', 10)
worksheet3.set_column('D:D', 25, format3)
worksheet3.set_column('E:E', 25)
worksheet3.set_column('F:F', 25)
worksheet3.set_column('G:G', 25)
worksheet3.set_column('H:H', 25, format1)
worksheet3.set_column('I:I', 25)
worksheet3.set_column('J:J', 25)
worksheet3.set_column('K:K', 25)
worksheet3.set_column('L:L', 25)
worksheet3.set_column('M:M', 25)
worksheet3.set_column('N:N', 25)
worksheet3.set_column('O:O', 25)
worksheet3.set_column('P:P', 25, format3)
worksheet3.set_column('Q:Q', 25, format2)
worksheet3.set_column('R:R', 25)
worksheet3.set_column('S:S', 25)
worksheet3.set_column('T:T', 25)
worksheet3.set_column('U:U', 25, format3)
worksheet3.set_column('V:V', 25)
worksheet3.set_column('W:W', 15)

#Formula 
worksheet3.write_dynamic_array_formula('W2','=REPT(0,1)&(RIGHT(H2,9))')
worksheet3.write_dynamic_array_formula('N2','=_xlfn.XLOOKUP($F2,LPC!$B:$B,LPC!$G:$G)')
worksheet3.write_dynamic_array_formula('O2','=_xlfn.XLOOKUP($F2,LPC!$B:$B,LPC!$H:$H)')
worksheet3.write_dynamic_array_formula('P2','=_xlfn.XLOOKUP($F2,PTP!$C:$C,PTP!$E:$E)')
worksheet3.write_dynamic_array_formula('Q2','=_xlfn.XLOOKUP($F2,PTP!$C:$C,PTP!$F:$F)')
worksheet3.write_dynamic_array_formula('R2','=_xlfn.XLOOKUP($F2,LPC!$B:$B,LPC!$L:$L)')
worksheet3.write_dynamic_array_formula('S2','=_xlfn.XLOOKUP($F2,LPC!$B:$B,LPC!$M:$M)')
worksheet3.write_dynamic_array_formula('T2',"=_xlfn.INDEX('Z:\\MIS\\Fone Wasin\\Python\\Call Activities\\[Seamoney Outbond Phone.xlsx]Sheet1'!$A:$A,RANDBETWEEN(1,COUNTA('Z:\\MIS\Fone Wasin\\Python\\Call Activities\\[Seamoney Outbond Phone.xlsx]Sheet1'!$A:$A)),1)")
worksheet3.write_dynamic_array_formula('U2','=Today()-1')
worksheet3.write_dynamic_array_formula('V2','=_xlfn.XLOOKUP($F2,LPC!$B:$B,LPC!$K:$K)')

print("Setting Format...is DONE!")
print("Saved " + todaysdate_filename)
print("Waitong For Open File: " + todaysdate_filename)
writer.save()

print(todaysdate_filename + " is DONE!")
# Open file or folder on OS
path_url = "Z:\\MIS\\Fone Wasin\\Python\\Call Activities\\"
path_file = path_url + "\*SEasyCash-BCL Daily report as of*.xls*"
for filex in glob.iglob(path_file, recursive=True):
    os.path.realpath(path_url)
    # FBI OPEN UP!!!!
    # os.startfile(path_url)
    os.startfile(filex)
    print('Opened File&Folder:', filex)
    
path_url = "C:\\Users\\wasin.k\\OneDrive - COLLECTIUS SYSTEMS PTE. LTD\\Desktop"
path_file = path_url + "\*Command MS*.xls*"
for filex in glob.iglob(path_file, recursive=True):
    os.path.realpath(path_url)
    # FBI OPEN UP!!!!
    # os.startfile(path_url)
    os.startfile(filex)
    print('Opened File&Folder:', filex)