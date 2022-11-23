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
import pyodbc
import xlsxwriter
from matplotlib.pyplot import axis
import zipfile

# ###MEX
# path_MEX = r'Z:\\MIS\\Fone Wasin\\Python\\GRAB SVC TH\\MEX\\'
# zip_MEX = path_MEX + '\*.zip'
# for zip_filename in glob.iglob(zip_MEX, recursive=True):
#     zip_handler = zipfile.ZipFile(zip_filename, "r") 
#     zip_handler.extractall(path_MEX)
#     print('UnZipfile: ' + zip_filename + ' is DONE!')

data_file_folder = "C:\\Users\\wasin.k\\OneDrive - COLLECTIUS SYSTEMS PTE. LTD\\GRAB\Payment\\MEX_Payment\\"

df = []
for file in os.listdir(data_file_folder):
    if file.endswith('.xlsx'):
        print('Loading file {0}...'.format(file))
        df.append(pd.read_excel(os.path.join(data_file_folder, file), sheet_name='ECA_MEX_TH_Daily'))
        
# Len(df)
df_combine = pd.concat(df, axis=0)
#df_combine2 = df_combine.iloc[:,[0,1,7,11]]
df_combine = df_combine[['report_date', 'debt_id', 'debtor_id', 'last_payment']]
df_combine = df_combine.assign(debt_id_text="",
                               Account_Number="",
                               Card_Number="",
                               Description="",
                               Amount="",
                               Amount_Amount_in_LCY="",
                               Effective_transaction_date="",
                               Transaction_Date_Posting="=TODAY()"
                               )

todaysdate_filename = str(
    datetime.now().strftime("CombineMEX'%Y%m%d")) + '.xlsx'
writer = pd.ExcelWriter(todaysdate_filename)
print("\n",df_combine, f"{todaysdate_filename }""\n")

df_combine.to_excel(writer, index=False, sheet_name= 'Combine_MEX')
# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['Combine_MEX']


# Add some cell formats.
format1 = workbook.add_format({'num_format': '@'})
format2 = workbook.add_format({'num_format': '0.00'})
format3 = workbook.add_format({'num_format': 'mm/dd/yyyy'})

# Set the column width and format.
worksheet.set_column('A:A', 12)
worksheet.set_column('B:B', 12)
worksheet.set_column('C:C', 15)
worksheet.set_column('D:D', 15, format2)
worksheet.set_column('E:E', 15)
worksheet.set_column('F:F', 30)
worksheet.set_column('G:G', 30)
worksheet.set_column('H:H', 14)
worksheet.set_column('I:I', 14, format2)
worksheet.set_column('J:J', 28, format2)
worksheet.set_column('K:K', 28, format3)
worksheet.set_column('L:L', 28, format3)
worksheet.set_column('M:M', 28, format3)


#Formula 
worksheet.write_dynamic_array_formula('E2', '=B2&""')
worksheet.write_dynamic_array_formula('F2', '=_xlfn.XLOOKUP(B2,Maping_MEX.xlsx!$A:$A,Maping_MEX.xlsx!$B:$B)')
worksheet.write_dynamic_array_formula('G2', '=_xlfn.XLOOKUP(B2,Maping_MEX.xlsx!$A:$A,Maping_MEX.xlsx!$B:$B)')
worksheet.write_dynamic_array_formula('I2', '=D2*1')
worksheet.write_dynamic_array_formula('J2', '=D2*1')
worksheet.write_dynamic_array_formula('K2', '=_xlfn.DATE(_xlfn.RIGHT(A2,4),_xlfn.MID(A2,4,2),_xlfn.LEFT(A2,2))')



# Close the Pandas Excel writer and output the Excel file.
writer.save()

month = str(datetime.now().strftime('%m-%y\\'))
# Move file on os base name and path
src_folder = r"C:\\Users\\wasin.k\\OneDrive - COLLECTIUS SYSTEMS PTE. LTD\\GRAB\Payment\\MEX_Payment\\"
dst_folder = r"C:\\Users\\wasin.k\\OneDrive - COLLECTIUS SYSTEMS PTE. LTD\\GRAB\Payment\\MEX_Payment\\Uploaded\\" 
target_folder = dst_folder + month
# move file whose name end with string 'xls'
pattern = src_folder + "\*.xls*"
for files in glob.iglob(pattern, recursive=True):
    # extract file name form file path
    file_name = os.path.basename(files)
    #todayy = str(datetime.now().strftime("(Uploaded) %H%M "))
    shutil.move(files, target_folder + file_name)
    print('Moved:', files)


path_url = r"Z:\\MIS\\Fone Wasin\\Python\\GRAB SVC TH\\"
# Open file or folder on OS
path_file2 = path_url + "\\*Maping_MEX*.xls*"
for filex in glob.iglob(path_file2, recursive=True):
    os.path.realpath(path_url)
    # FBI OPEN UP!!!!
   #os.startfile(path_url)
    os.startfile(filex)
    print('Opened File&Folder:', filex)  
    break

# Open file or folder on OS
path_file = path_url + "\\*CombineMEX*.xls*"
for filex in glob.iglob(path_file, recursive=True):
    os.path.realpath(path_url)
    # FBI OPEN UP!!!!
   #os.startfile(path_url)
    os.startfile(filex)
    print('Opened File&Folder:', filex)
    break
