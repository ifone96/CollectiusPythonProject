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

data_file_folder = "Z:\\MIS\\Fone Wasin\\Python\\Clean DATA\\DATA\\"

df = []
for file in os.listdir(data_file_folder):
    if file.endswith('.xlsx'):
        print('Loading file {0}...'.format(file))
        df.append(pd.read_excel(os.path.join(
            data_file_folder, file), sheet_name='All Port'))

# Len(df)
df_combine = pd.concat(df, axis=0)


df_combine1 = df_combine.loc[:, ['เบอร์โทรศัพท์']]
df_combine2 = df_combine1.loc[:, ['ECA', 'AGREEMENT NO',
                                'Call date', 'Contact disposition', 'Call out come',
                                'Description', 'Call from', 'Portflorio']].astype('string')
reCol = {
    'ECA': 'subject',
    'AGREEMENT': 'account number',
    'เบอร์โทรศัพท์': 'phone number',
    'วันที่ติดตามล่าสุด': 'contact dispoisition',
    'Portflorio': 'portfolio'
}

# call rename () method
df_combine2.rename(columns=reCol, inplace=True)
# add column
df_combine2.insert(2, "invoice number", "=B2", True)

todaysdate_filename = str(
    datetime.datetime.now().strftime("Test %H%M")) + '.xlsx'
writer = pd.ExcelWriter(todaysdate_filename)
print("\n", df_combine2, f"{todaysdate_filename }""\n")

df_combine2.to_excel(writer, index=False, sheet_name='Sheet1')

workbook = writer.book
worksheet = writer.sheets['Sheet1']


# Add some cell formats.
format1 = workbook.add_format({'num_format': '@'})
format2 = workbook.add_format({'num_format': '0.00'})
format3 = workbook.add_format({'num_format': 'm/dd/yy'})


# Set the column width and format.
worksheet.set_column('A:A', 20)
worksheet.set_column('B:B', 20)
worksheet.set_column('C:C', 20)
worksheet.set_column('D:D', 20,format2)
worksheet.set_column('E:E', 20)
worksheet.set_column('F:F', 20)
worksheet.set_column('G:G', 20)
worksheet.set_column('H:H', 20)
worksheet.set_column('I:I', 20)
worksheet.set_column('J:J', 20)
worksheet.set_column('K:K', 20)
worksheet.set_column('L:L', 20)
worksheet.set_column('M:M', 20)


# Close the Pandas Excel writer and output the Excel file.
writer.save()

# Move file on os base name and path
#src_folder = r"C:\\Users\\wasin.k\\Desktop\\Python\\Run PD\\SMN\\From TL\\"
#dst_folder = r"C:\\Users\\wasin.k\\Desktop\\Python\\Run PD\\SMN\\From TL\\Uploaded\\"
# move file whose name end with string 'xls'
#pattern = src_folder + "\*.xls*"
# for files in glob.iglob(pattern, recursive=True):
#    # extract file name form file path
#    file_name = os.path.basename(files)
#    #todayy = str(datetime.datetime.now().strftime("(Uploaded) %H%M "))
#    shutil.move(files, dst_folder + file_name)
#    print('Moved:', files)

# Open file or folder on OS
path_url = r"Z:\\MIS\\Fone Wasin\\Python\\Clean DATA\\"
path_file = path_url + "\*.xls*"
for filex in glob.iglob(path_file, recursive=True):
    os.path.realpath(path_url)
    # FBI OPEN UP!!!!
    os.startfile(path_url)
    os.startfile(filex)
    print('Opened File&Folder:', filex)
