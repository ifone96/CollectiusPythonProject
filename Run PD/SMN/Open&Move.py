from doctest import DocFileTest
from email.utils import format_datetime
from math import fabs
from operator import index
import os
from pickle import NONE
from tkinter import W
from matplotlib.pyplot import axis
import pandas as pd
import datetime
import xlsxwriter
import uuid
import pyodbc
import shutil
import glob


data_file_folder = "C:\\Users\\wasin.k\\Desktop\\Python\\Run PD\\SMN\\From TL"

df = []
for file in os.listdir(data_file_folder):
    if file.endswith('.xlsb'):
        print('Loading file Name: {0}'.format(file))
        df.append(pd.read_excel(os.path.join(data_file_folder, file), engine='pyxlsb',sheet_name='Sheet1'))
    if file.endswith('.xlsx'):
        print('Loading file Name: {0}'.format(file))
        df.append(pd.read_excel(os.path.join(data_file_folder, file), sheet_name='Sheet1'))

#Move file on os base name and path
src_folder = r"C:\\Users\\wasin.k\\Desktop\\Python\\Run PD\\SMN\\From TL\\"
dst_folder = r"C:\\Users\\wasin.k\\Desktop\\Python\\Run PD\\SMN\\From TL\\Uploaded\\"

# move file whose name end with string 'xls'
pattern = src_folder + "\*.xls*"
for files in glob.iglob(pattern, recursive=True):
    # extract file name form file path
    file_name = os.path.basename(files)
    todayy = str(datetime.datetime.now().strftime("(Uploaded) %H%M "))
    shutil.move(files, dst_folder + todayy + file_name)
    print('Moved:', files)

#Open file or folder on OS
path_url = r"C:\\Users\\wasin.k\\Desktop\\Python\\Run PD\\SMN\\"
path_file = path_url + "\*.xls*"
for filex in glob.iglob(path_file, recursive=True):
    os.path.realpath(path_url)
    #FBI OPEN UP!!!!
    os.startfile(path_url)
    os.startfile(filex)
    print('Opened File&Folder:', filex)