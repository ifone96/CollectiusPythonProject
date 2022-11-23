import win32com.client 
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
from email.parser import Parser

outlook=win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI") 
inbox=outlook.GetDefaultFolder(6) #Inbox default index value is 6 
message =inbox.Items
message2= message.GetLast()
subject=message2.Subject
body=message2.body 
date=message2.senton.date()    
sender=message2.Sender 
attachments=message2.Attachments 
print(subject) 
print(body) 
print(sender) 
print(attachments.count) 
print(date)  
