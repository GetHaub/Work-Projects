import pandas as pd
from pandas.api.types import is_datetime64_any_dtype as is_datetime
import numpy as np
import os
import glob
import time
import datetime
import math
import win32com.client as win32
from win32com.client import Dispatch
from pywintypes import com_error
import traceback
import SCM_Ora as SOF
import pyodbc
# import comtypes, comtypes.client



def send_outlook_email(recipients, sender='SLIRheem@schneider.com', subject='No Subject', body='Blank', send_or_display='Display', copies=None, 
                       attach=None, template_filepath=None):
    """
    Parameters
    ----------
    recipients : List
        List of recipients' email addresses
    sender : str
        str type, since there can only be 1 sender! Email address of sender
    subject : str
        Subject line of email
    body : str
        Body of email
    send_or_display : str
        If value is "Send" then the email automatically sends, otherwise the email window pops up
    copies : List
        List of cc addresses
    attach : str
        File path to attachment file, if necessary
    template_filepath : str
        File path to Outlook template file, if sending a template
    
    Returns
    -------
    None
    """
    print('Running send_outlook_email function...')
    if len(recipients) > 0:
    # if len(recipients) > 0 and isinstance(recipient_list, list):
        # outlook = win32.client.Dispatch("Outlook.Application")
        outlook = win32.Dispatch("Outlook.Application")
        
        if template_filepath == None:
            mail = outlook.CreateItem(0)
            str_to = ""
            for recipient in recipients:
                str_to += recipient + ";"
            mail.To = str_to
            
            if copies is not None:
                str_cc = ""
                for cc in copies:
                    str_cc += cc + ";"
                mail.CC = str_cc
            
            mail.Subject = subject
            mail.Body = body
            mail.SentOnBehalfOfName = sender # Will all outbound emails be sent from this address?
            if attach != None:
                mail.Attachments.Add(attach)
            
            if send_or_display.upper() == 'SEND':
                mail.Send()
            else:
                mail.Display()
        else:  # Just using the provided Outlook template
            mail = outlook.CreateItemFromTemplate(template_filepath)
            str_to = ""
            for recipient in recipients:
                str_to += recipient + ";"
            mail.To = str_to
            mail.Subject = subject
            mail.SentOnBehalfOfName = sender
            if attach != None:
                mail.Attachments.Add(attach)
            mail.GetInspector
            if send_or_display.upper() == 'SEND':
                mail.Send()
            else:
                mail.Display()
    else:
        print('Recipient email address - NOT FOUND')


def check_email_reply(reply_subject, sender):
    print('Running check_email_reply function...')
    try:
        # reply_subject = "RE: Status request for all loads scheduled for delivery today for SCAC code " + str(scac)
        # reply_subject = 'RE: M302015776'
        outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
        # accounts = Dispatch("Outlook.Application").Session.Accounts
        rheem_index = 0
        print('rheem_index')
        for i in range(outlook.Folders.Count):
            # print(outlook.Folders[i])
            if str(outlook.Folders[i]) == 'SLI Rheem':
                rheem_index = i
        rheem_inbox_index = 0
        print('rheem_inbox_index')
        for i in range(outlook.Folders[rheem_index].Folders.Count):
            # print(outlook.Folders[rheem_index].Folders[i])
            if str(outlook.Folders[rheem_index].Folders[i]) == 'Inbox':
                rheem_inbox_index = i
        
        inbox = outlook.Folders[rheem_index].Folders[rheem_inbox_index]
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True) # Sorts in order of time received
        today = datetime.date.today()
        last_hour = datetime.datetime.now() - datetime.timedelta(hours = 1) # Check if response was in the last hour
        # last_hour_emails = messages.Restrict("[ReceivedTime] >= '" + last_hour.strftime('%m/%d/%Y %H:%M %p')+"'")
        
        print('\n')
        count = 0
        for msg in messages:
            if msg.subject == reply_subject and msg.senton.date() == today and msg.senton.time() >= last_hour.time() and msg.SenderEmailAddress == sender:
                print(msg.SenderEmailAddress)
                print(type(msg.SenderEmailAddress))
                count += 1
                print(count)
                did_respond = True
        if count == 1:
            did_respond = True
            print(msg.SenderEmailAddress)
            print('The provided subject and sender email address were found in the inbox, received within the past hour')
        elif count >= 1:
            did_respond = True
            print(msg.SenderEmailAddress)
            print('MULTIPLE EMAILS: The provided subject and sender email address were found in the inbox more than once, all received within the past hour')
        else:
            did_respond = False
            print('Not Found: The provided subject and sender email address were NOT found in the inbox, within the past hour')
    except Exception as e:
        print('Exception raised! %s \n' % e)
        body_email = "Exception raised during execution of check_email_reply function. Full traceback below. \n\n" + str(traceback.format_exc())
        send_outlook_email(recipients=['TangE1@schneider.com', 'sliinfomgmt@schneider.com'], sender='TangE1@schneider.com', subject='Exception raised in function check_email_reply within Rheem_Automation.py', 
                           body=body_email, send_or_display='Send')
    return did_respond


def download_reply_email_attachment(scac):
    print('Running retrieve_response function...')
    try:
        reply_subject = "RE: Status request for all loads scheduled for delivery today for SCAC code " + str(scac)
        outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
        # accounts = Dispatch("Outlook.Application").Session.Accounts
        rheem_index = 0
        print('rheem_index')
        for i in range(outlook.Folders.Count):
            # print(outlook.Folders[i])
            if str(outlook.Folders[i]) == 'SLI Rheem':
                rheem_index = i
        rheem_inbox_index = 0
        print('rheem_inbox_index')
        for i in range(outlook.Folders[rheem_index].Folders.Count):
            # print(outlook.Folders[rheem_index].Folders[i])
            if str(outlook.Folders[rheem_index].Folders[i]) == 'Inbox':
                rheem_inbox_index = i
        
        inbox = outlook.Folders[rheem_index].Folders[rheem_inbox_index]
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True) # Sorts in order of time received
        today = datetime.date.today()
        last_hour = datetime.datetime.now() - datetime.timedelta(hours = 1)
        # last_hour_emails = messages.Restrict("[ReceivedTime] >= '" + last_hour.strftime('%m/%d/%Y %H:%M %p')+"'")
        
        # Download the attachment
        for msg in messages:
            if msg.subject == reply_subject and msg.senton.date() == today and msg.senton.time() >= last_hour.time():
                attachments = msg.attachments
                print(attachments.count)
                # for i in range(attachments.count):
                attachment = attachments.Item(1) # Each response email should only have 1 attachment
                attachment.SaveAsFile(os.path.join(r'C:\Users\w95997\Desktop\Rheem\Testing', str(attachment)))
    except Exception as e:
        print('Exception raised! %s \n' % e)
        body_email = "Exception raised during execution of download_reply_email_attachment function. Full traceback below. \n\n" + str(traceback.format_exc())
        send_outlook_email(recipients=['TangE1@schneider.com', 'sliinfomgmt@schneider.com'], sender='TangE1@schneider.com', subject='Exception raised in function download_reply_email_attachment within Rheem_Automation.py', 
                           body=body_email, send_or_display='Send')


def download_rheem_reports(year, month, day):
    # Download the RHM Daily Shipments Report from the SLI Rheem inbox
    print('Running download_rheem_reports function...')
    try:
        daily_reports_subject_line = 'RHM Daily Shipments Report ' + year + month + day
        acd_folder = r'C:\Users\w95997\Desktop\Rheem\RHM Daily Shipments Report\ACD'
        whd_folder = r'C:\Users\w95997\Desktop\Rheem\RHM Daily Shipments Report\WHD'
        outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
        accounts = Dispatch("Outlook.Application").Session.Accounts
        rheem_index = 0
        print('rheem_index')
        for i in range(outlook.Folders.Count):
            # print(outlook.Folders[i])
            if str(outlook.Folders[i]) == 'SLI Rheem':
                rheem_index = i
        rheem_inbox_index = 0
        print('rheem_inbox_index')
        for i in range(outlook.Folders[rheem_index].Folders.Count):
            # print(outlook.Folders[rheem_index].Folders[i])
            if str(outlook.Folders[rheem_index].Folders[i]) == 'Inbox':
                rheem_inbox_index = i
        pro_update_index = 0
        print('pro_update_index')
        for i in range(outlook.Folders[rheem_index].Folders[rheem_inbox_index].Folders.Count):
            print(outlook.Folders[rheem_index].Folders[rheem_inbox_index].Folders[i])
            if str(outlook.Folders[rheem_index].Folders[rheem_inbox_index].Folders[i]) == '*PRO # Update':
                pro_update_index = i
        
        inbox = outlook.Folders[rheem_index].Folders[rheem_inbox_index]
        messages_inbox = inbox.Items
        pro_update_inbox = outlook.Folders[rheem_index].Folders[rheem_inbox_index].Folders[pro_update_index]
        messages_pro_update = pro_update_inbox.Items
        # for i in range(outlook.Folders[rheem_index].Folders[pro_update_index].Folders.Count):
        #     print(outlook.Folders[rheem_index].Folders[pro_update_index].Folders[i], i)
        
        x = 0 # The point of this counter is so that the 2 daily attachments are saved in 2 different folders (ACD & WHD)
        for msg in messages_pro_update:
            # print(msg)
            if msg.subject == daily_reports_subject_line and msg.Senton.date() == datetime.date.today():
                attachments = msg.attachments
                if x == 0:
                    print('Retrieving first daily report')
                    for i in range(attachments.count):
                        attachment = attachments.Item(i + 1)
                        print(str(attachment))
                        attachment_filepath = os.path.join(acd_folder, str(attachment))
                        if os.path.exists(attachment_filepath) == False:
                            attachment.SaveAsFile(attachment_filepath)
                        else:
                            print('Another file already exists in the ACD folder with the same name of the file that you are trying to save. Using existing file.')
                if x == 1:
                    print('Retrieving second daily report')
                    for i in range(attachments.count):
                        attachment = attachments.Item(i + 1)
                        print(str(attachment))
                        attachment_filepath = os.path.join(whd_folder, str(attachment))
                        if os.path.exists(attachment_filepath) == False:
                            attachment.SaveAsFile(attachment_filepath)
                        else:
                            print('Another file already exists in the WHD folder with the same name of the file that you are trying to save. Using existing file.')
            # else:
            #     print('Did not find the daily reports subject for today.')
                x += 1
        
        # If the daily shipments report isn't in the *PRO # Update inbox folder, check the inbox
        if (os.path.exists(r'C:\Users\w95997\Desktop\Rheem\RHM Daily Shipments Report\ACD\XXONT_DAILY_SHIPMENTS_' + year + '_' + month + '_' + day + '.xls') == False) & (os.path.exists(r'C:\Users\w95997\Desktop\Rheem\RHM Daily Shipments Report\WHD\XXONT_DAILY_SHIPMENTS_' + year + month + day + '.xls') == False):
            y = 0 # The point of this counter is so that the 2 daily attachments are saved in 2 different folders (ACD & WHD)
            for msg in messages_inbox:
                # print(msg)
                if msg.subject == daily_reports_subject_line and msg.Senton.date() == datetime.date.today():
                    print('2')
                    attachments = msg.attachments
                    if y == 0:
                        print('y == 0')
                        for i in range(attachments.count):
                            attachment = attachments.Item(i + 1)
                            print(str(attachment))
                            attachment_filepath = os.path.join(acd_folder, str(attachment))
                            if os.path.exists(attachment_filepath) == False:
                                attachment.SaveAsFile(attachment_filepath)
                            else:
                                print('Another file already exists in the ACD folder with the same name of the file that you are trying to save. Using existing file.')
                    if y == 1:
                        print('y == 1')
                        for i in range(attachments.count):
                            attachment = attachments.Item(i + 1)
                            print(str(attachment))
                            attachment_filepath = os.path.join(whd_folder, str(attachment))
                            if os.path.exists(attachment_filepath) == False:
                                attachment.SaveAsFile(attachment_filepath)
                            else:
                                print('Another file already exists in the WHD folder with the same name of the file that you are trying to save. Using existing file.')
                # else:
                #     print('Did not find the daily reports subject for today.')
                    y += 1
    except Exception as e:
        print('Exception raised! %s \n' % e)
        body_email = "Exception raised during execution of download_rheem_reports function. Full traceback below. \n\n" + str(traceback.format_exc())
        send_outlook_email(recipients=['TangE1@schneider.com', 'sliinfomgmt@schneider.com'], sender='TangE1@schneider.com', subject='Exception raised in function download_rheem_reports within Rheem_Automation.py', 
                           body=body_email, send_or_display='Send')


def format_df(df, ev=True, email_template=False):
    """
    For data that is not consistent with the formatting of NL Tracking.xlsx, convert 
    that data into strings and str format it to look like the formatting in NL Tracking
    
    Parameters
    ----------
    df : pandas dataframe
        Unformatted dataframe
    ev : boolean
        If true, the passed in df also has EV columns at the end
    email_template : boolean
        Email templates include columns PO# up to and including Comments
    
    Returns
    -------
    df : pandas dataframe
        Formatted dataframe, good to export to Excel
    """
    # pandas to_excel exports NaN and NaT values as blank cells in Excel
    # Format dates/times/datetimes to be consistent with existing data
    # Actual Time column code (commented out) causes an error
    # name =[x for x in globals() if globals()[x] is df][0] # str name of df so that it can be printed
    # print('Running format_df function for dataframe', name, '...') # Print the name of the df
    print('Running format_df function...')
    # df = future_ev.copy()
    # df.drop([600,601,602,603], axis=0, inplace=True)
    try:
        df['PO#'] = df['PO#'].astype(str)
        df['Loaded Date'] = df['Loaded Date'].astype(str)
        df['Loaded Date'] = pd.to_datetime(df['Loaded Date']).dt.strftime('%m/%d/%Y')
        df['Delivery Appointment'] = df['Delivery Appointment'].astype(str)
        df['Delivery Appointment'] = pd.to_datetime(df['Delivery Appointment']).dt.strftime('%m/%d/%Y')
        df['Delivery Time'] = df['Delivery Time'].astype(str)
        df['New Appointment'] = df['New Appointment'].astype(str)
        df['Actual Time'] = df['Actual Time'].astype(str)
        if ev == True and email_template == False:  # Email attachments won't have EV so no need for & ev == True
            df['Dept Req'] = df['Dept Req'].astype(str)
            df['Last EV Location Update Datetime'] = df['Last EV Location Update Datetime'].astype(str)
            df['Last EV ETA Update Datetime'] = df['Last EV ETA Update Datetime'].astype(str)
            i = 0
            while i < len(df):
                if df.at[i, 'Delivery Time'] != 'nan':
                    df.at[i, 'Delivery Time'] = df.at[i, 'Delivery Time'][:5] # Remove the seconds portion of the times
                if (df.at[i, 'New Appointment'] != 'NaT') & (df.at[i, 'New Appointment'].upper() != 'NAN') & (df.at[i, 'New Appointment'].upper() != 'PENDING'):
                    df.at[i, 'New Appointment'] = df.at[i, 'New Appointment'][5:7] + '/' + df.at[i, 'New Appointment'][8:10] + '/' + df.at[i, 'New Appointment'][:4] + df.at[i, 'New Appointment'][10:16]
                # if (df.at[i, 'Actual Time'] != 'NaT') & (df.at[i, 'Dept Req'].upper() != 'NAN'):
                #     df.at[i, 'Actual Time'] = df.at[i, 'Actual Time'].str[:5]
                if (df.at[i, 'Dept Req'] != 'NaT') & (df.at[i, 'Dept Req'].upper() != 'NAN'):
                    df.at[i, 'Dept Req'] = df.at[i, 'Dept Req'][5:7] + '/' + df.at[i, 'Dept Req'][8:10] + '/' + df.at[i, 'Dept Req'][:4]
                if df.at[i, 'Last EV Location Update Datetime'] != 'NaT':
                    df.at[i, 'Last EV Location Update Datetime'] = df.at[i, 'Last EV Location Update Datetime'][5:7] + '/' + df.at[i, 'Last EV Location Update Datetime'][8:10] + '/' + df.at[i, 'Last EV Location Update Datetime'][:4] + df.at[i, 'Last EV Location Update Datetime'][10:16]
                if df.at[i, 'Last EV ETA Update Datetime'] != 'NaT':
                    df.at[i, 'Last EV ETA Update Datetime'] = df.at[i, 'Last EV ETA Update Datetime'][5:7] + '/' + df.at[i, 'Last EV ETA Update Datetime'][8:10] + '/' + df.at[i, 'Last EV ETA Update Datetime'][:4] + df.at[i, 'Last EV ETA Update Datetime'][10:16]
                i += 1
        elif ev == True and email_template == True:
            print('EMAIL TEMPLATE EXCEL FILES SHOULD NOT BE INCLUDING EV FIELDS!')
        elif ev == False and email_template == True:
            i = 0
            while i < len(df):
                if df.at[i, 'Delivery Time'] != 'nan':
                    df.at[i, 'Delivery Time'] = df.at[i, 'Delivery Time'][:5] # Remove the seconds portion of the times
                if (df.at[i, 'New Appointment'] != 'NaT') & (df.at[i, 'New Appointment'].upper() != 'NAN') & (df.at[i, 'New Appointment'].upper() != 'PENDING'):
                    df.at[i, 'New Appointment'] = df.at[i, 'New Appointment'][5:7] + '/' + df.at[i, 'New Appointment'][8:10] + '/' + df.at[i, 'New Appointment'][:4] + df.at[i, 'New Appointment'][10:16]
                # if (df.at[i, 'Actual Time'] != 'NaT') & (df.at[i, 'Dept Req'].upper() != 'NAN'):
                #     df.at[i, 'Actual Time'] = df.at[i, 'Actual Time'].str[:5]
                i += 1
        else:  # ev == False & email_template == False
            df['Dept Req'] = df['Dept Req'].astype(str)
            i = 0
            while i < len(df):
                if df.at[i, 'Delivery Time'] != 'nan':
                    df.at[i, 'Delivery Time'] = df.at[i, 'Delivery Time'][:5] # Remove the seconds portion of the times
                if (df.at[i, 'New Appointment'] != 'NaT') & (df.at[i, 'New Appointment'].upper() != 'NAN') & (df.at[i, 'New Appointment'].upper() != 'PENDING'):
                    df.at[i, 'New Appointment'] = df.at[i, 'New Appointment'][5:7] + '/' + df.at[i, 'New Appointment'][8:10] + '/' + df.at[i, 'New Appointment'][:4] + df.at[i, 'New Appointment'][10:16]
                # if (df.at[i, 'Actual Time'] != 'NaT') & (df.at[i, 'Dept Req'].upper() != 'NAN'):
                #     df.at[i, 'Actual Time'] = df.at[i, 'Actual Time'].str[:5]
                if (df.at[i, 'Dept Req'] != 'NaT') & (df.at[i, 'Dept Req'].upper() != 'NAN'):
                    df.at[i, 'Dept Req'] = df.at[i, 'Dept Req'][5:7] + '/' + df.at[i, 'Dept Req'][8:10] + '/' + df.at[i, 'Dept Req'][:4]
                i += 1
        df.replace('NaT', '', inplace=True)
        df.replace('nan', np.nan, inplace=True)
    # return df
    except Exception as e:
        print('Exception raised! %s \n' % e)
        body_email = "Exception raised during execution of format_df function. Full traceback below. \n\n" + str(traceback.format_exc())
        send_outlook_email(recipients=['TangE1@schneider.com', 'sliinfomgmt@schneider.com'], sender='TangE1@schneider.com', subject='Exception raised in function format_df within Rheem_Automation.py', 
                           body=body_email, send_or_display='Send')
    finally:
        return df


def export_fourkites_files(filepath, export=True):
    """
    Use execute_macro() function [Add_EV_Fields macro] right after to copy the FourKites data into the NL Tracking.xlsx file itself
    If all goes well, the following block should only need to be executed once.
    Create the Excel files with the NL Tracking data merged with the EV (FourKites) data 
    so that the VBA can use those files as the reference tables when filling out the vlookups
    
    Parameters
    ----------
    filepath : str
        File path to the target NL Tracking file path that you want to read and populate (via VBA) with FourKites data
    export : boolean
        If True, export Excel files for each tab (Delivered, Deliver Today, Future) 
        to use to paste into NL Tracking
    
    Returns
    -------
    None
    """
    # filepath = r'\\dom1\sfs\Shared\Logistic\Team Docs\Rheem\Tracking\NL\NL Tracking FourKites.xlsx'
    # export = True
    print('Running export_fourkites_files function...')
    try:
        hstp0 = (r'C:\Users\w95997\FDUS_Engi\hstp0.ora')
        attributes_sql = r'\\dom1\sfs\Shared\Logistic\Shared Docs\BIA\Projects\Rheem NL Dray Tracking\SQL\ATTRIBUTES.sql'
        print('Running SOF.get_data function, called within export_fourkites_files function')
        attributes_time = time.time()
        attributes = SOF.get_data(hstp0, attributes_sql)
        print('SOF.get_data function took %s minute(s) and %d second(s)' 
              % (math.floor((time.time() - attributes_time) / 60), (time.time() - attributes_time) % 60), '\n')
        attributes.rename({'SHIPMENT_XID': 'Load'}, axis=1, inplace=True)
        attributes.rename({'ATTRIBUTE10': 'Last EV Location Update', 
                    'ATTRIBUTE11': 'Last EV Tracking Status Update', 
                    'ATTRIBUTE12': 'Last EV ETA StopReferenceID Update', 
                    'ATTRIBUTE_DATE1': 'Last EV Location Update Datetime', 
                    'ATTRIBUTE_DATE3': 'Last EV ETA Update Datetime'}, axis=1, inplace=True)
        # Add in the EV (FourKites) columns to be exported to the reference files that VBA will use to copy and paste to NL Tracking
        print('Reading the Delivered worksheet into pandas')
        delivered = pd.read_excel(filepath, sheet_name='Delivered')
        delivered_ev = pd.merge(delivered, attributes, how='left', on='Load')
        print('Reading the Deliver Today worksheet into pandas')
        deliver_today = pd.read_excel(filepath, sheet_name='Deliver Today')
        deliver_today_ev = pd.merge(deliver_today, attributes, how='left', on='Load')
        print('Reading the Future worksheet into pandas \n')
        future = pd.read_excel(filepath, sheet_name='Future')
        future_ev = pd.merge(future, attributes, how='left', on='Load')
        if export == True:
            # Exports to Excel folder in shared drive
            # format_df(delivered_ev).to_excel(r'\\dom1\sfs\Shared\Logistic\Team Docs\Rheem\Tracking\NL\Automation\Excel\Delivered.xlsx', index=False)
            # format_df(deliver_today_ev).to_excel(r'\\dom1\sfs\Shared\Logistic\Team Docs\Rheem\Tracking\NL\Automation\Excel\Deliver Today.xlsx', index=False)
            # format_df(future_ev).to_excel(r'\\dom1\sfs\Shared\Logistic\Team Docs\Rheem\Tracking\NL\Automation\Excel\Future.xlsx', index=False)
            # format_df(delivered_ev).to_excel(r'C:\Users\w95997\Desktop\Rheem\Delivered.xlsx', index=False)
            # format_df(deliver_today_ev).to_excel(r'C:\Users\w95997\Desktop\Rheem\Deliver Today.xlsx', index=False)
            # format_df(future_ev).to_excel(r'C:\Users\w95997\Desktop\Rheem\Future.xlsx', index=False)
            delivered_ev.to_excel(r'C:\Users\w95997\Desktop\Rheem\Delivered.xlsx', index=False)
            deliver_today_ev.to_excel(r'C:\Users\w95997\Desktop\Rheem\Deliver Today.xlsx', index=False)
            future_ev.to_excel(r'C:\Users\w95997\Desktop\Rheem\Future.xlsx', index=False)
    except Exception as e:
        print('Exception raised! %s \n' % e)
        body_email = "Exception raised during execution of export_fourkites_files function. Full traceback below. \n\n" + str(traceback.format_exc())
        send_outlook_email(recipients=['TangE1@schneider.com', 'sliinfomgmt@schneider.com'], sender='TangE1@schneider.com', subject='Exception raised in function export_fourkites_files within Rheem_Automation.py', 
                           body=body_email, send_or_display='Send')


def execute_macro(xlsm_filepath, macro_name):
    if os.path.exists(xlsm_filepath):
        # print('Executing macro:', str(macro_name))
        # excel_app = win32.Dispatch('Excel.Application')
        # excel_app = win32.dynamic.Dispatch('Excel.Application')
        excel_app = win32.DispatchEx('Excel.Application')
        # print('DispatchEx has executed')
        excel_app.DisplayAlerts = False
        excel_app.Visible = True
        wb = excel_app.Workbooks.Open(Filename=xlsm_filepath)
        print('Executing macro:', str(macro_name))
        try:
            excel_app.Application.Run(macro_name)
        except com_error as e:
            print(str(macro_name), 'macro has generated an error!')
            body_email = "Exception raised during execution of execute_macro function. Full traceback below. \n\n" + str(traceback.format_exc())
            send_outlook_email(recipients=['TangE1@schneider.com', 'sliinfomgmt@schneider.com'], sender='TangE1@schneider.com', subject='Exception raised in function execute_macro within Rheem_Automation.py', 
                               body=body_email, send_or_display='Send')
        finally:
            wb.Close() # Don't have vba close the last excel file
            excel_app.Application.Quit()
            print('If no error was generated,', str(macro_name), 'has finished running')
    else:
        print('The provided file path does not exist')


