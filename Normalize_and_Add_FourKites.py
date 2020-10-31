import pandas as pd
import numpy as np
import os
import time
import datetime
import math
import multiprocessing
import win32com.client as win32
from win32com.client import Dispatch
from shutil import copyfile
import SCM_Ora as SOF
import Rheem_Automation as rheem

# This script integrates the EV fields into the regular NL Tracking.xlsx file, and 
# saves as NL Tracking FourKites.xlsx

start_time = time.time()
pd.set_option('display.max_columns', None) # Shows all columns

vba_folderpath = r'\\dom1\sfs\Shared\Logistic\Team Docs\Rheem\Tracking\NL\Automation'
vba_tracking_file = os.path.join(vba_folderpath, 'NL TRACKING AUTOMATION.xlsm') # The xlsm file containing the macros

today = datetime.date.today()
current_year = str(today.year)
if today.month < 10:
    current_month = '0' + str(today.month)
else:
    current_month = str(today.month)
if today.day < 10:
    current_day = '0' + str(today.day)
else:
    current_day = str(today.day)

# ---------------------------------------------------

# Copy NL Tracking.xlsx into the Automation folder
try:
    copyfile(src=r'\\dom1\sfs\Shared\Logistic\Team Docs\Rheem\Tracking\NL\NL Tracking.xlsx', dst=r'C:\Users\w95997\Desktop\Rheem\NL Tracking FourKites.xlsx')
except Exception as e:
    print('Could not copy source file to destination file! \nReason: %s' % e)

# Save the NL Tracking.xlsx as NL Tracking FourKites.xlsx first
rheem.execute_macro(vba_tracking_file, 'Normalize_Data') # Executes on NL Tracking FourKites.xlsx

# Create separate exported Excel files for each tab, for VBA to use to populate NL Tracking.xlsx
# rheem.export_fourkites_files(r'C:\Users\w95997\Desktop\Rheem\Excel\NL Tracking FourKites.xlsx', export=True)
rheem.export_fourkites_files(r'C:\Users\w95997\Desktop\Rheem\NL Tracking FourKites.xlsx', export=True)
# Populate FourKites data in NL Tracking.xlsx with the exported Excels
rheem.execute_macro(xlsm_filepath=vba_tracking_file, macro_name='Add_EV_Fields')
rheem.execute_macro(vba_tracking_file, 'Normalize_Data')

try:
    # nl_tracking_fourkites = open(r'\\dom1\sfs\Shared\Logistic\Team Docs\Rheem\Tracking\NL\NL Tracking FourKites.xlsx', 'r+')
    copyfile(src=r'C:\Users\w95997\Desktop\Rheem\NL Tracking FourKites.xlsx', dst=r'\\dom1\sfs\Shared\Logistic\Team Docs\Rheem\Tracking\NL\NL Tracking FourKites.xlsx')
except Exception as e:
    print('Could not copy source file to destination file! \nReason: %s' % e)


print('\nRuntime took %s minute(s) and %d second(s)' 
      % (math.floor((time.time() - start_time) / 60), (time.time() - start_time) % 60), '\n')

time.sleep(60) # Pause program completion for 60 seconds so you can see the program run time from the line above