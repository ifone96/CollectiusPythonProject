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

src_folder = r"Z:\\MIS\\Fone Wasin\\Python\\Call Activities\\"
dst_folder = r"Z:\\MIS\\Fone Wasin\\Python\\Call Activities\\Done\\"
# move file whose name end with string 'xls'
pattern = src_folder + "*Daily report as of*.xls*"
for files in glob.iglob(pattern, recursive=True):
    # extract file name form file path
    file_name = os.path.basename(files)
    #todayy = str(datetime.now().strftime("(Uploaded) %H%M "))
    shutil.move(files, dst_folder + file_name)
    print('Moved:', files)


