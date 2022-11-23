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
import time

start = time.time()
print("Start: "+ str(start))


data_file_folder = "Z:\\MIS\\Fone Wasin\\Python\\Run PD\\SMN\\From TL\\"
df = []
for file in os.listdir(data_file_folder):
    if file.endswith('.xlsb'):
        print('Loading file name: {0}'.format(file))
        df.append(pd.read_excel(os.path.join(data_file_folder, file), engine='pyxlsb', sheet_name='Sheet1'))
    if file.endswith('.xlsx'):
        print('Loading file name: {0}'.format(file))
        df.append(pd.read_excel(os.path.join(data_file_folder, file), sheet_name='Sheet1'))
        
#combine
df_combine = pd.concat(df, axis='index')
# Pick up column with headername
df_combine = df_combine[['bill_id']].astype('string')
print('Getting values: bill_id')
#df_combine = df_combine.drop_duplicates()
df_combine = df_combine.assign(uuid="",
                               phone="",
                               type="",
                               name="",
                               idnumber="")
print('Assign values: uuid phone type name idnumber')

# server = 'localhost\sqlexpress' # for a named instance
# server = 'myserver,port' # to specify an alternate port
server = 'collectiusdwhph.database.windows.net'
database = 'dwh_th_2022'
username = 'atiwat'
password = '2a#$dfERat^%'
print('Connecting database: '+ server + '...')
connect_database = pyodbc.connect(
    'DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD=' + password)
querying_s = time.time()
print('SQL querying...'+ str(querying_s))
sql_cmd = """
    SELECT DISTINCT
    b.alternis_invoicenumber as bill_id, 
    b.alternis_accountid as uuid,
    REPLACE(a.alternis_number,'*','') AS phone,
    a.alternis_phonetypename as type,
    a.alternis_contactidname as name,
    b.alternis_idnumber as idnumber
    FROM stage.alternis_phone a
    JOIN stage.alternis_account b
    ON a.alternis_contactid = b.alternis_contactid
    WHERE alternis_portfolioidname IN ('SEAMONEY SPL SVC TH','SEAMONEY BCL SVC TH')
    """
df_sql = pd.read_sql(sql_cmd, connect_database)
querying_e = time.time()
querying_x = querying_e - querying_s
print('SQL query done...'+ str(querying_x))

print('Waiting for Mapping with XLOOKUP method...')
# f(x) xlookup python pandas
def xlookup(lookup_value, lookup_array, return_array, if_not_found: str = ''):
    match_value = return_array.loc[lookup_array == lookup_value]
    if match_value.empty:
        # return f'"{lookup_value}" is NULL' if if_not_found == '' else if_not_found
        return f'NULL' if if_not_found == '' else if_not_found
    else:
        return match_value.tolist()[0]   
    
#xlookup_data = xlookup('1625976231738529792', df_sql['bill_id'],df_sql['uuid'])
df_combine['uuid'] = df_combine['bill_id'].apply(
    xlookup, args=(df_sql['bill_id'], df_sql['uuid']))
df_combine['phone'] = df_combine['bill_id'].apply(
    xlookup, args=(df_sql['bill_id'], df_sql['phone']))
df_combine['type'] = df_combine['bill_id'].apply(
    xlookup, args=(df_sql['bill_id'], df_sql['type']))
df_combine['name'] = df_combine['bill_id'].apply(
    xlookup, args=(df_sql['bill_id'], df_sql['name']))
df_combine['idnumber'] = df_combine['bill_id'].apply(
    xlookup, args=(df_sql['bill_id'], df_sql['idnumber']))
print('XLOOKUP method... is DONE!')

# delete some u don't need
del df_combine['bill_id']
print('Deleted: bill_id')
#del df_sql['bill_id']

# Set name file with date/times
todaysdate_filename = str(
    datetime.datetime.now().strftime("Leads - SMN %H%M")) + '.xlsx'
writer = pd.ExcelWriter(todaysdate_filename)
print('Created file name: ' + todaysdate_filename)

#Write file bill_id
df_combine.to_excel(writer, index=False, engine='xlsxwriter', sheet_name='SMN_PD')
print('Created Sheet name: SMN_PD')
#Write File SQL First
#df_sql.to_excel(writer, index=False, sheet_name='SQL_MAP')
#print('Created Sheet name: SQL_MAP' + '\n''by Python ')

# Get the xlsxwriter workbook and worksheet objects.
workbook = writer.book
worksheet = writer.sheets['SMN_PD']
#worksheet2 = writer.sheets['SQL_MAP']

# Add some cell formats.
header_format = workbook.add_format({'bold': True})

# Set the column width and format.
worksheet.set_row(0, None, header_format)
worksheet.set_column('A:A', 40)
worksheet.set_column('B:B', 16)
worksheet.set_column('C:C', 20)
worksheet.set_column('D:D', 30)
worksheet.set_column('E:E', 25)
worksheet.set_column('F:F', 25)

#orksheet2.set_row(0, None, header_format)
#orksheet2.set_column('A:A', 25)
#orksheet2.set_column('B:B', 40)
#orksheet2.set_column('C:C', 16)
#orksheet2.set_column('D:D', 20)
#orksheet2.set_column('E:E', 30)
#orksheet2.set_column('F:F', 25)

# Close the Pandas Excel writer and output the Excel file.
writer.save()
print('Execute code with python is done!')

# Move file on os base name and path
src_folder = r"Z:\\MIS\\Fone Wasin\\Python\\Run PD\\SMN\\From TL\\"
dst_folder = r"Z:\\MIS\\Fone Wasin\\Python\\Run PD\\SMN\\From TL\\Uploaded\\"
# move file whose name end with string 'xls'
pattern = src_folder + "\*.xls*"
for files in glob.iglob(pattern, recursive=True):
    # extract file name form file path
    file_name = os.path.basename(files)
    #todayy = str(datetime.datetime.now().strftime("(Uploaded) %H%M "))
    shutil.move(files, dst_folder + file_name)
    print('Moved:', files)

# Open file or folder on OS
path_url = r"Z:\\MIS\\Fone Wasin\\Python\\Run PD\\SMN"
path_file = path_url + "\*.xls*"
for filex in glob.iglob(path_file, recursive=True):
    os.path.realpath(path_url)
    # FBI OPEN UP!!!!
    os.startfile(path_url)
    os.startfile(filex)
    print('Opened File&Folder:', filex)
    
end = time.time()
xxx = end-start
print("End: "+ str(xxx))

