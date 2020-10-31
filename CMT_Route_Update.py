from selenium import webdriver
import time
from datetime import datetime, timedelta
import pandas as pd
import os
#import snipy.util.config as cg #I AM LIKELLY MISSING CONFIG PACKAGE 
import SCM_Ora as SOF
from datetime import datetime
from selenium.webdriver.common.keys import Keys
#from pywinauto.keyboard import send_keys
import pyperclip

# get_username = input("Please provide your username: ")
# get_password = input("Please provide your password: ")


# get_username = "l78924"
# get_password = "176059Cj!!!"
get_username = "w95997"
get_password = "Graduated2018"

df_input = pd.read_excel(r"\\dom1\remote_shared\Farmington Hills\fcsd\Engineering"
                       "\Python Projects\CMT web\Files\dateupdatemaster.xlsx")

df_input = df_input[df_input["Y/N"]==1]

xdock_numpy = df_input['xdock'].unique()

xdock_list =[]
for each in xdock_numpy:
    xdock_list.append(each)
    
df_date = pd.read_excel(r"\\dom1\remote_shared\Farmington Hills\fcsd\Engineering"
                       "\Python Projects\CMT web\Files\Datevar.xlsx")

#start date that will be used for SQL
startdate = df_date['Date'].unique()
startdate = str(startdate)[2:12]
startdate = datetime.strptime(startdate, '%Y-%m-%d') 
startdate = datetime.strftime(startdate, '%d-%b-%Y').upper()

print('Generating SQL')
#delete all info from sql file that will be used
with open(r"\\dom1\remote_shared\Farmington Hills\fcsd\Engineering"
        "\Python Projects\CMT web\SQL\cmtroute2.sql","w")as delFile:
    delFile.write(' ')

sqldate = ("TO_CHAR(EST_XDCK_DTTM, 'DD-MON-YYYY HH24:MI') BETWEEN " + "'" +
                                               str(startdate).upper() + " 00:00" + "'" +' and '
                                               +"'"+
                                                  str(startdate).upper() + " 23:59" + "'"+  "\n")

sqlxdock = str(xdock_list)
sqlxdock = sqlxdock.replace('[', '(').replace(']', ')')
sqlxdock = "XDOCK_PTY_ID in " + sqlxdock                                              
groupby = "GROUP BY\n" + "XDOCK_PTY_ID,\n" + "ROUTE_SK_ID\n"    
      
#look at sql wihtout parameters
with open(r"\\dom1\remote_shared\Farmington Hills\fcsd\Engineering"
        "\Python Projects\cmt web\SQL\cmtroute1.sql","r") as f:
    f = str(f.read())
    
with open(r"\\dom1\remote_shared\Farmington Hills\fcsd\Engineering"
        "\Python Projects\cmt web\SQL\cmtroute2.sql","a+") as f2:
        #write sql with parameters into new sql to pull on next step
    f2.write(f + sqldate + "and\n" + sqlxdock + "\n" + groupby)

#sql that will be used to extract data from oracle     
sql = (r"\\dom1\remote_shared\Farmington Hills\fcsd\Engineering"
        "\Python Projects\cmt web\SQL\cmtroute2.sql")

#ora file filepath
oraFilepath = (r'c:\users' + '\\' + get_username + '\FDUS_Engi')

db = oraFilepath
print("db path is ", db)
print('Extracting data from Oracle hstp0')
hstp0 = os.path.join(db, 'hstp0.ora')
ecdw0 = os.path.join(db, 'ecdw.ora')
    
# query HSTP0 to get data
df_sql = SOF.get_data(hstp0, sql)

df_sql["xdock"]=df_sql["XDOCK_PTY_ID"]

df_final = pd.merge(df_input,df_sql)

#delete/comment out after test  
df_final = pd.DataFrame(data=[[3971658,4,1]],columns=["ROUTE_SK_ID","MAX_STOP","Days_Added"])

df_final = df_final[["ROUTE_SK_ID","MAX_STOP","Days_Added"]]

#df = pd.DataFrame({"ROUTE_SK_ID":"2","MAX_STOP":"s","Days_Added":"s"},index=[1])
#df_final = df_final.append(df)
##############################################################################

#get_username = 'l78924'
#get_password = '176059Cj!'
print("running selenium...")
for index,row in df_final.iterrows():
    
    driver = webdriver.Chrome()
    driver.get("http://sb1.intra.schneider.com/cmt/ui/faces/jsps/cmthomepage.jspx")
    driver.maximize_window()
    time.sleep(4)
    
    input_username = driver.find_element_by_css_selector("#username")
    input_password = driver.find_element_by_css_selector("#password")
    
    input_username.send_keys(get_username)
    input_password.send_keys(get_password)
    
    driver.find_element_by_css_selector("#submit-button > a").click()
    time.sleep(4)
                                        
    driver.find_element_by_css_selector("#pt1\:r1\:0\:sdi4\:\:disAcr ").click()                                   
    time.sleep(3)  
    #x=2
                         
    driver.find_element_by_css_selector("#pt1\:r1\:0\:r3\:0\:qryId1\:val00\:\:content").clear()
                                 
    input_routeid = driver.find_element_by_css_selector("#pt1\:r1\:0\:r3\:0\:qryId1\:val00\:\:content")
    input_routeid.send_keys(str(row["ROUTE_SK_ID"]),Keys.ENTER )                                            
    #input_routeid.send_keys("3760453",Keys.ENTER)  
    time.sleep(5)
  
    driver.find_element_by_css_selector("#pt1\:r1\:0\:r3\:0\:pc1\:routeTable\:0\:ot14").click()  
    time.sleep(5) 
    
    copied_text = driver.find_element_by_css_selector("#pt1\:r1\:0\:r3\:1\:id2\:\:content").get_attribute("value")     #innerHTML                     
    
    datetime_object = datetime.strptime(copied_text, '%m-%d-%Y %H:%M:%S') 
    datetime_object =  datetime_object + timedelta(days=int(row["Days_Added"]))   
    #datetime_object =  datetime_object + timedelta(days=x)                                             
    datetime_object = datetime_object.strftime("%m-%d-%Y %H:%M:%S")     
                                                   
    driver.find_element_by_css_selector("#pt1\:r1\:0\:r3\:1\:id2\:\:content").clear()
                                        
    input_date_value = driver.find_element_by_css_selector("#pt1\:r1\:0\:r3\:1\:id2\:\:content")
    
    input_date_value.send_keys(str(datetime_object))    
                                     
    i=0                                                              
    while int(row["MAX_STOP"]) > i :
    #while 4 > i :
            
        try:  
            input_stopsequence_value = driver.find_element_by_css_selector("#pt1\:r1\:0\:r3\:1\:pc1\:t1\:_afrFltrMdlc10\:\:content")
            driver.find_element_by_css_selector("#pt1\:r1\:0\:r3\:1\:pc1\:t1\:_afrFltrMdlc10\:\:content").clear()
            input_stopsequence_value.send_keys(str(i+1),Keys.ENTER)
            time.sleep(5)
                
            pyperclip.copy(' ')
                 
            #Schedule Arrival Date/Time date change
            copied_textA = driver.find_element_by_css_selector("#pt1\:r1\:0\:r3\:1\:pc1\:t1\:" + str(i) + "\:id4\:\:content").get_attribute("value") 
            datetime_objectA = datetime.strptime(copied_textA, '%m-%d-%Y %H:%M:%S') 
            datetime_objectA = datetime_objectA + timedelta(days=int(row["Days_Added"]))    
            #datetime_objectA = datetime_objectA + timedelta(days=x)  
            datetime_objectA = datetime_objectA.strftime("%m-%d-%Y %H:%M:%S") 
            driver.find_element_by_css_selector("#pt1\:r1\:0\:r3\:1\:pc1\:t1\:" + str(i) + "\:id4\:\:content").clear()
            input_date_valueA = driver.find_element_by_css_selector("#pt1\:r1\:0\:r3\:1\:pc1\:t1\:" + str(i) + "\:id4\:\:content")
            input_date_valueA.send_keys(str(datetime_objectA)[:10], Keys.CONTROL,"v")  
            input_date_valueA.send_keys(str(datetime_objectA)[10:])
            
            #Schedule Departure Date/Time date change                                                
            copied_textB = driver.find_element_by_css_selector("#pt1\:r1\:0\:r3\:1\:pc1\:t1\:" + str(i) + "\:id6\:\:content").get_attribute("value")
            datetime_objectB = datetime.strptime(copied_textB, '%m-%d-%Y %H:%M:%S') 
            datetime_objectB = datetime_objectB + timedelta(days=int(row["Days_Added"]))   
            #datetime_objectB = datetime_objectB + timedelta(days=x)                                   
            datetime_objectB = datetime_objectB.strftime("%m-%d-%Y %H:%M:%S")                                                        
            driver.find_element_by_css_selector("#pt1\:r1\:0\:r3\:1\:pc1\:t1\:" + str(i) + "\:id6\:\:content").clear()                            
            input_date_valueB = driver.find_element_by_css_selector("#pt1\:r1\:0\:r3\:1\:pc1\:t1\:" + str(i) + "\:id6\:\:content")
            input_date_valueB.send_keys(str(datetime_objectB)[:10], Keys.CONTROL,"v")  
            input_date_valueB.send_keys(str(datetime_objectB)[10:])
               
            i+=1
        except:
            pass
        
    #driver.find_element_by_css_selector("#pt1\:r1\:0\:r3\:1\:cb1").click()
    time.sleep(5)
    #driver.find_element_by_css_selector("#pt1\:r1\:0\:r3\:1\:cb3").click()
    time.sleep(5)
    driver.close()