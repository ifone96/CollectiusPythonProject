from ast import If
import datetime
from datetime import datetime, timedelta
import glob
import os
import shutil
from tkinter import HIDDEN
from unittest import skip
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

# Move file on os base name and path
src_folder = r"Z:\\MIS\\Fone Wasin\\Python\\Payment Plan Daily\\"
dst_folder = r"Z:\\MIS\\Fone Wasin\\Python\\Payment Plan Daily\\Uploads Result\\"
# move file whose name end with string 'xls'
pattern = src_folder + "\*OA*.xls*"
for files in glob.iglob(pattern, recursive=True):
    # extract file name form file path
    file_name = os.path.basename(files)
    #todayy = str(datetime.datetime.now().strftime("(Uploaded) %H%M "))
    shutil.move(files, dst_folder + file_name)
    print('Moved:', files)
    break