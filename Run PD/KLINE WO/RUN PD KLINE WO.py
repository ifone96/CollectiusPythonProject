from doctest import DocFileTest
from email.utils import format_datetime
from operator import index
import os
from matplotlib.pyplot import axis
import pandas as pd
import datetime
import xlsxwriter

data_file_folder = "C:\\Users\\wasin.k\\Desktop\\Python\\Run PD\\KLINE WO\\"

df = []
for file in os.listdir(data_file_folder):
    if file.endswith('.xlsx'):
        print('Loading file {0}...'.format(file))
        df.append(pd.read_excel(os.path.join(data_file_folder, file), sheet_name='Sheet1'))
        
# Len(df)
df_combine = pd.concat(df, axis=0)
#df_combine2 = df_combine.iloc[:,[0,1,7,11]]
#Pick up column with headernamer
df_combine2 = df_combine[['uuid', 'Phone Number', 'Phone Type', 'Debtor Name']]
#add column
df_combine3 = df_combine2.assign(idnumber ='NULL')
#rename column
df_combine4 = df_combine3.rename(columns={"Phone Number": "phone", "Phone Type": "type" , "Debtor Name": "name"})
#fill NaN or Blank set to value
Nullrow = {"phone": "NULL", "type": "NULL" , "name": "NULL"}
df_combine5 = df_combine4.fillna(value=Nullrow)
#Set name file with date/times
todays_date_name = str(datetime.datetime.now().strftime("KLINE WO %H%M") )+ '.xlsx'
writer = pd.ExcelWriter(todays_date_name, engine = 'xlsxwriter')
df_combine5.to_excel(writer, index=False, sheet_name= 'KLINE WO')

# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['KLINE WO']

# Add some cell formats.
format1 = workbook.add_format({'num_format': '@'})
format2 = workbook.add_format({'num_format': '0000000000'})

# Set the column width and format.
worksheet.set_column('A:A', 40, format1)
worksheet.set_column('B:B', 16, format2)
worksheet.set_column('C:C', 10, format1)
worksheet.set_column('D:D', 22, format1)
worksheet.set_column('E:E', 20, format1)


# Close the Pandas Excel writer and output the Excel file.
writer.save()