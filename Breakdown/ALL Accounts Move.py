#All Accounts Move
import zipfile
from ast import Delete, If
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

src_folder = r"C:\\Users\\wasin.k\\OneDrive - COLLECTIUS SYSTEMS PTE. LTD\\Documents - MIS-TH\\Morning Data\\"
dst_folder = r"C:\\Users\\wasin.k\\OneDrive - COLLECTIUS SYSTEMS PTE. LTD\\Documents - MIS-TH\\Morning Data\\Older Date Files\\"
# move file whose name end with string 'xls'
pattern = src_folder + "*All*.xls*"
for files in glob.iglob(pattern, recursive=True):
    # extract file name form file path
    file_name = os.path.basename(files)
    #todayy = str(datetime.now().strftime("(Uploaded) %H%M "))
    shutil.move(files, dst_folder + file_name)
    print('Moved:', files)

src_folder = "Z:\\MIS\\Fone Wasin\\Python\\Breakdown\\"
dst_folder = "C:\\Users\\wasin.k\\OneDrive - COLLECTIUS SYSTEMS PTE. LTD\\Documents - MIS-TH\\Morning Data"
pattern = src_folder + "\\*All*.xls*"
for files in glob.iglob(pattern, recursive=True):
    file_name = os.path.basename(files)
    shutil.copy(files, dst_folder) 
    print('Moved:', files)

src_folder = r"Z:\\MIS\\Fone Wasin\\Python\\Breakdown\\"
dst_folder = r"Z:\\MIS\\Fone Wasin\\Python\\Breakdown\\Uploaded\\" 
# move file whose name end with string 'xls'
pattern = src_folder + "\*All*.xls*"
for files in glob.iglob(pattern, recursive=True):
    # extract file name form file path
    file_name = os.path.basename(files)
    #todayy = str(datetime.now().strftime("(Uploaded) %H%M "))
    shutil.move(files, dst_folder + file_name)
    print('Moved:', files)
    
path_url = "C:\\Users\\wasin.k\\OneDrive - COLLECTIUS SYSTEMS PTE. LTD\\Desktop"
path_file = path_url + "\*Command MS*.xls*"
for filex in glob.iglob(path_file, recursive=True):
    os.path.realpath(path_url)
    # FBI OPEN UP!!!!
    # os.startfile(path_url)
    os.startfile(filex)
    print('Opened File&Folder:', filex)
