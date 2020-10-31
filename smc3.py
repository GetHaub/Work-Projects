import CA3_Functions
from selenium import webdriver
import time
from selenium.webdriver.support.ui import Select
import zipfile
import os 
import pandas as pd
import glob
import datetime
import win32com.client as win32

#if you get this error for line 23
#AttributeError: module 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9' 
#has no attribute 'CLSIDToPackageMap'

#Do the following
#navigate to "C:\Users\<your username>\AppData\Local\Temp\gen_py" and 
#somewhere within that directory you will find a folder with the same name as 
#displayed in the AttributeError message (and based on the original post, 
#the op should delete the folder titled '00020813-0000-0000-C000-000000000046x0x1x9')

def ReSave(pivot_table_csv):
    excel = win32.gencache.EnsureDispatch('Excel.Application') # opens Excel
    wb = excel.Workbooks.Open(pivot_table_csv) 
    wb.Save()
    wb.Close()
    excel.Quit()

#variables to log in
def smc3_report(email,batchmarkpwd,user,pivot_table_csv):
    #name= 'milesC@schneider.com'
    #pwd = 'mFr3rdY6'
    #schneider_Id = 'l78924'
    today = datetime.date.today()
    
    #go to website
    driver = webdriver.Chrome()
    driver.get("https://batchmarkxl.smc3.com/BatchMarkXL/")
    
    #log in function
    CA3_Functions.login(email,batchmarkpwd,driver)
    
    #go to map input tab and wait
    driver.get("https://batchmarkxl.smc3.com/BatchMarkXL/FileMap/mapInput.smc")
    time.sleep(5)
    
    #File paths
    file = (pivot_table_csv)
    
    downloaded_file = (r'\\dom1\remote_shared\Farmington Hills\WasteQuip\Reporting'
                       '\Python Projects - WAST\LTL_COST_SAVINGS\Downloaded Files')
    output_file = (r'\\dom1\remote_shared\Farmington Hills\WasteQuip\Reporting'
                            '\Python Projects - WAST\LTL_COST_SAVINGS\Output Files')
    
    #uploading file function
    CA3_Functions.uploadFile(driver,'map-file-input',file)
    
    # map_file = pd.read_csv(r"\\dom1\remote_shared\Farmington Hills\WasteQuip\Reporting"
    #                     "\Python Projects - WAST\LTL_COST_SAVINGS\Output Files\Wastequip_Pivot_Table.csv")
    
    # headers = list(map_file)[-1]
    
    FAK_class = "WQ_Multi_2_FAK"
    
    #need to check the current FAK CLASS everytime this file is ran
    CA3_Functions.selectInListbox(driver, 'fileMapId',FAK_class)
    
    #wait for file to load
    time.sleep(30)
    
    
    #Code to know which drop down menu items to update  
    
    
    #update drop down menus for specific FAK Class
    select = Select(driver.find_element_by_id('sel_11'))
    select.select_by_visible_text('Sum of TOTAL_CLASS_WEIGHT (FAK Class Rank = 1)')
    
    select = Select(driver.find_element_by_id('sel_33'))
    select.select_by_visible_text('Sum of TOTAL_CLASS_WEIGHT (FAK Class Rank = 2)')
    
    select = Select(driver.find_element_by_id('sel_32'))
    select.select_by_visible_text('Average of FAK Class Structure for Rating (FAK Class Rank = 2)')
    
    #submit application
    driver.find_element_by_css_selector('#file-map-submit-btn').click()
    
    #wait on email 
    print('Getting results through email')
    
    def outlook_is_running():
        import win32ui
        try:
            win32ui.FindWindow(None, "Microsoft Outlook")
            return True
        except win32ui.error:
            return False
    
    if not outlook_is_running():
        os.startfile("outlook")
    
    time.sleep(40)
    
    link = CA3_Functions.getEmails(email)
    
    driver.get(link)
    
    #wait until link has downloaded 
    time.sleep(30)
    
    driver.close()
    
    #go to downloads folder and get most recent zip file
    downloads_folder = glob.glob("C:\\Users\\" + user + "\\" + "Downloads\*.zip")
    latest_file = max(downloads_folder, key=os.path.getctime)
    
    #open zip file , extract file and save in downloaded folder and close zip file
    zip_ref = zipfile.ZipFile(latest_file)
    zip_ref.extractall(downloaded_file) 
    zip_ref.close()
     
    #grab files from downloaded files folder and get length # of files
    smcResults = os.listdir(downloaded_file),'\\' '*final.csv'
    first_len_list = len(smcResults[0])
    
    #get first file in downloaded files folder 
    os.rename(downloaded_file + '\\' + smcResults[0][first_len_list-1],
              downloaded_file + '\\' + 'SMC3_Results.csv')
    
    #remove zip file from downloads folder
    os.remove(latest_file)
    
    #read csv and remove first 3 rows
    smc3File = pd.read_csv(downloaded_file + '\\' + 'SMC3_Results.csv'
                           ,skiprows = [0,1,2])
    
    os.remove(downloaded_file + '\\' + 'SMC3_Results.csv')
    smc3File.to_csv(output_file + '\\' + str(today) +' SMC3_Results.csv',index=False)
    
    smc3_filepath = output_file + '\\' + str(today) +' SMC3_Results.csv'
    
    return(smc3_filepath)

def refresh_final_report(report):
    if report == 'LTL':
        final_report = (r'\\dom1\remote_shared\Farmington Hills\WasteQuip\Reporting'
                          '\Python Projects - WAST\LTL_COST_SAVINGS'
                          '\Output Files\Final_Report.xlsm')
        
    elif (report == 'TRUCKLOAD') | (report == 'TL'):
        final_report = (r'\\dom1\remote_shared\Farmington Hills\WasteQuip\Reporting'
                          '\Python Projects - WAST\LTL_COST_SAVINGS'
                          '\Output Files\Final_Report_TL.xlsm')
      
    if os.path.exists(final_report):
        xlApp = win32.Dispatch('Excel.Application')
        xlApp.DisplayAlerts = False
        #xlApp.Visible = True
        wb = xlApp.Workbooks.Open(Filename=final_report)
        xlApp.Application.Run('Macro1')
        wb.Close() #dont have vba close the last excel file
        xlApp.Application.Quit()
    
