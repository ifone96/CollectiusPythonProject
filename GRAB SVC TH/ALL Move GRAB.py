#Move GRAB
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

month = str(datetime.now().strftime('%m-%y\\'))
# Move file on os base name and path
src_folder = r"Z:\\MIS\\Fone Wasin\\Python\\GRAB SVC TH\\"
dst_folder = r"Z:\\MIS\\Fone Wasin\\Python\\GRAB SVC TH\\Uploaded Payment\\"
target_folder = dst_folder
# move file whose name end with string 'xls'
pattern = src_folder + "\*Combine*X*.xls*"
for files in glob.iglob(pattern, recursive=True):
    # extract file name form file path
    file_name = os.path.basename(files)
    #todayy = str(datetime.now().strftime("(Uploaded) %H%M "))
    shutil.move(files, target_folder + file_name)
    print('Moved:', files)

# # DELECTE MEX
# for MexZip in glob.iglob("Z:\\MIS\\Fone Wasin\\Python\\GRAB SVC TH\\MEX\\*.zip"):
#     os.remove(MexZip)
#     break
# # DELECTE DAX
# for DaxZip in glob.iglob("Z:\\MIS\\Fone Wasin\\Python\\GRAB SVC TH\\DAX\\*.zip"):
#     os.remove(DaxZip)
#     break
# print('Delete Files Zip is DONE!')

path_url = "C:\\Users\\wasin.k\\OneDrive - COLLECTIUS SYSTEMS PTE. LTD\\Desktop"
path_file = path_url + "\*Command MS*.xls*"
for filex in glob.iglob(path_file, recursive=True):
    os.path.realpath(path_url)
    # FBI OPEN UP!!!!
    # os.startfile(path_url)
    os.startfile(filex)
    print('Opened File&Folder:', filex)
    break
