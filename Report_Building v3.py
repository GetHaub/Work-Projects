import os
import numpy as np
import pandas as pd
import math
import SCM_Ora as SOF
import time
import datetime
import smc3

report = input('Which report would you like to run? ').upper()
user = input('Please enter your user ID: ')
email = input('Please enter your schneider email: ')
batchmarkpwd = input('Please enter your batchmark profile password: ')


inputs_folder = (r"\\dom1\remote_shared\Farmington Hills\Wastequip\Reporting"
                   "\Python Projects - WAST\LTL_COST_SAVINGS\Input Files")

output_folder = (r"\\dom1\remote_shared\Farmington Hills\Wastequip\Reporting"
                   "\Python Projects - WAST\LTL_COST_SAVINGS\Output Files")

sql_folder = (r"\\dom1\remote_shared\Farmington Hills\Wastequip\Reporting"
                   "\Python Projects - WAST\LTL_COST_SAVINGS\SQL")

def Oracle_Queries(user, report, sql_folder):
    db = (r'C:\Users\\' + user + '\\FDUS_Engi')
    ecdw0 = os.path.join(db, 'ecdw.ora')
    hstp0 = os.path.join(db, 'hstp0.ora')
    
    #retreiving data from Oracle
    print("Extracting data from Oracle")
    if report == 'LTL':
        #hstp = SOF.get_data(hstp0, sql_folder + '\\' + 'HSTP0_sql_query_as_shipmentgid.sql') # used to be C: drive
        ecdw = SOF.get_data(ecdw0, sql_folder + '\\' + 'wasteecdw23_as_shipmentgid.sql') # used to be C: drive
        hstp_new = SOF.get_data(hstp0, sql_folder + '\\' + 'hstp0_new_as_shipmentgid.sql') # used to be C: drive

    elif (report == 'TRUCKLOAD') | (report == 'TL'):
        #hstp_TL = SOF.get_data(hstp0, sql_folder + '\\' + 'HSTP0_sql_query_as_shipmentgid_TL.sql') # used to be C: drive
        ecdw_TL = SOF.get_data(ecdw0, sql_folder + '\\' + 'wasteecdw23_as_shipmentgid_TL.sql') # used to be C: drive
        hstp_new_TL = SOF.get_data(hstp0, sql_folder + '\\' + 'hstp0_new_as_shipmentgid_TL.sql') # used to be C: drive
        
        #hstp_FB = SOF.get_data(hstp0, sql_folder + '\\' + 'HSTP0_sql_query_as_shipmentgid_FB.sql') # used to be C: drive
        ecdw_FB = SOF.get_data(ecdw0, sql_folder + '\\' + 'wasteecdw23_as_shipmentgid_FB.sql') # used to be C: drive
        hstp_new_FB = SOF.get_data(hstp0, sql_folder + '\\' + 'hstp0_new_as_shipmentgid_FB.sql') # used to be C: drive
     
        #hstp = hstp_TL.append(hstp_FB, ignore_index=True)
        ecdw = ecdw_TL.append(ecdw_FB, ignore_index=True)
        hstp_new = hstp_new_TL.append(hstp_new_FB, ignore_index=True)
        
    audit_rate = SOF.get_data(hstp0, sql_folder + '\\' + 'otm_audit_rate.sql')
        
    return(ecdw, hstp_new, audit_rate) 
    
def Aggregating_Data_Part_One(ecdw, hstp_new, inputs_folder):
    # today = datetime.date.today()
    
    #Modifying and aggregating data
    accounting_periods = pd.read_csv(inputs_folder + '\\' + 'AccountingPeriods.csv')
    accounting_periods['Period Start'] = pd.to_datetime(accounting_periods['Period Start'])
    accounting_periods['Period End'] = pd.to_datetime(accounting_periods['Period End'])
    
    acct_periods = pd.read_csv(inputs_folder + '\\' + 'AccountingPeriodByDay.csv')
    acct_periods['Date'] = pd.to_datetime(acct_periods['Date']).apply(lambda x: x.date())
    acct_periods['Date (Calc)'] = pd.to_datetime(acct_periods['Date (Calc)']).apply(lambda x: x.date())
    

    cost_lbs = pd.read_csv(inputs_folder + '\\' + 'Cost_Lbs.csv', dtype={'Cost Unit (by Plant)': str})
    cost_lbs['Division'] = cost_lbs['Division'].replace('WASTEQUIP STEEL DIVISION', 'STEEL')
    cost_lbs['Commonality'] = cost_lbs['Division'] + cost_lbs['Cost Unit (by Plant)'].astype(str) + cost_lbs['Segment']
    
    # Changed from inner join to left join
    merge_new = pd.merge(hstp_new, ecdw, how='left', on='SHIPMENT_GID')
    
    mapping_file = open(inputs_folder + '\\' + "CUST_FRT_CAT_VAL_Mapping.txt","r")
    for line in mapping_file:
        line = line.strip().split(',')
        merge_new = merge_new.replace(line[0],line[1])
        
    merge_new['Cost_Unit_by_Plant'] = merge_new['Cost_Unit_by_Plant'].str.replace('000', '00')
    merge_new['Commonality'] = merge_new['DIVISION'] + merge_new['Cost_Unit_by_Plant'].astype(str) + merge_new['CUST_FRT_CAT_VAL']
    
    dfa = pd.merge(cost_lbs, merge_new, how='right', on='Commonality')
    
    #remove this in the future
    dfa['Avg Historical Miles'] = dfa['Average Miles'] #48
    dfa['Avg Historic Weight'] = dfa['Average Weight'] #49
    
    dfa['Stop Alt WT'] = '19999'
    dfa['PKP_PLN_DPRT_DTTM'] = dfa['PKP_PLN_DPRT_DTTM'].dt.date # Critical for later code
    
    # Populate the columns Wastequip Fiscal Period, Week Number, Week Start, and Week End
    # PKP_PLN_DPRT_DTTM
    dfa['Week Number'] = 0 # There will be no 0 values after everything is done
    dfa['Week Start'] = ''
    dfa['Week End'] = ''
    
    i = 0
    while i < len(dfa):
        j = 0
        while j < len(accounting_periods):
            if (dfa.at[i, 'PKP_PLN_DPRT_DTTM'] >= accounting_periods.at[j, 'Period Start']) & (dfa.at[i, 'PKP_PLN_DPRT_DTTM'] <= accounting_periods.at[j, 'Period End']):
                dfa.at[i, 'Wastequip Fiscal Period'] = accounting_periods.at[j, 'Accounting Period']
                break
            j += 1
        pkp = dfa.at[i, 'PKP_PLN_DPRT_DTTM']
        week_index = acct_periods[acct_periods['Date'] == pkp].index.item()
        week_num = acct_periods.at[week_index, 'Week #']
        dfa.at[i, 'Week Number'] = week_num
        tempdf = acct_periods[acct_periods['Week #'] == week_num].copy()
        tempdf.reset_index(inplace=True)
        dfa.at[i, 'Week Start'] = tempdf.Date.min() # This assumption holds since the dates are sorted in ascending order
        dfa.at[i, 'Week End'] = tempdf.Date.max()
        i += 1
    
    del_columns = open(inputs_folder + '\\' + "Unwanted_Columns.txt","r")
    for line in del_columns:
        line = line.strip('\n')
        dfa.drop([line], axis = 1, inplace=True) 
    
    dfa['EQUIPMENT_GROUP_GID'] = dfa['EQUIPMENT_GROUP_GID'].str[4:]
    
    return dfa

def nmfc(dfa, inputs_folder):
    
    freight_classes = pd.read_excel(inputs_folder + '\\' + 'Freight Class Types.xlsx')
    freight_classes['FAK'] = freight_classes['FAK'].astype(str).str.replace('.0', '', regex=False)
    freight_classes['WQ Standard'] = freight_classes['WQ Standard'].astype(str).str.replace('.0', '', regex=False)
    freight_classes['WQ Exception'] = freight_classes['WQ Exception'].astype(str).str.replace('.0', '', regex=False)
    
    dfa['NMFC_CLASS Converted'] = dfa['NMFC_CLASS'].str.replace('.0', '', regex=False) # String; removes trailing .0's
    
    i = 0
    while i < len(dfa):
        # FAK Class Structure for Rating
        j = freight_classes[freight_classes['FAK'] == dfa.at[i, 'NMFC_CLASS Converted']].index.item() # This works because all values in the FAK column in freight_classes are unique
        if dfa.at[i, 'ZIP3_ORG'] == '788': # Then use WQ Exception
            dfa.at[i, 'FAK Class Structure for Rating'] = freight_classes.at[j, 'WQ Exception'] # string
        else:
            dfa.at[i, 'FAK Class Structure for Rating'] = freight_classes.at[j, 'WQ Standard'] # string
        i += 1
    
    dfa['FAK Class Rank'] = 0 # After everything is done, there will be no 0 values
    
    # Determine FAK Class Rank, which considers the FAK Class Structure for Rating column 
    i = 0
    while i < len(dfa):
        sid = dfa.at[i, 'SHIPMENT_GID']
        rating = float(dfa.at[i, 'FAK Class Structure for Rating']) # String
        subset = dfa[dfa['SHIPMENT_GID'] == sid].copy()
        subset['FAK Class Structure for Rating'] = subset['FAK Class Structure for Rating'].astype(float)
        subset.sort_values(by=['FAK Class Structure for Rating'], inplace=True) # GENERATES A SETTING WITH COPY WARNING !!! Line ~151
        ratings_sorted = list(subset['FAK Class Structure for Rating']) # list of floats
        ratings_uniques = []
        # Removes duplicates from list so that an accurate FAK CLass Rank can be retrieved
        # [ratings_uniques.append(x) for x in ratings_sorted if x not in ratings_uniques]
        for r in ratings_sorted:
            if r not in ratings_uniques:
                ratings_uniques.append(r)
        ratings_uniques.sort()
        dfa.at[i, 'FAK Class Rank'] = ratings_uniques.index(rating) + 1
        i += 1
    
    # Convert columns BW and BX into formatted strings so that non-integer values will show as is and integer values won't have a decimal with a trailing zero
    dfa['FAK Class Structure for Rating'] = dfa['FAK Class Structure for Rating'].str.replace('.0', '', regex=False) # Already string values
    dfa['FAK Class Structure for Rating'] = dfa['FAK Class Structure for Rating'].astype(int)
    
    return(dfa)

def audit_rate_dataframes(audit_rate):
    # Incurs SettingWithCopyWarnings
    otm_ahc = audit_rate[(audit_rate['SP_SCAC'] == "ZHC1")]
    otm_ahc.reset_index(inplace=True)
    otm_ebc = audit_rate[(audit_rate['SP_SCAC'] == "ZBC1")]
    
    otm_ahc['Lane'] = (otm_ahc['S_ST'].astype(str) + ' ' + otm_ahc['S_REGION_ID'].astype(str).str[-3:] 
                       + ':' + otm_ahc['D_ST'].astype(str) + ' ' + otm_ahc['D_REGION_ID'].astype(str).str[-3:])
    otm_ebc['Lane'] = (otm_ebc['S_ST'].astype(str) + ' ' + otm_ebc['S_REGION_ID'].astype(str).str[-3:] 
                       + ':' + otm_ebc['D_ST'].astype(str) + ' ' + otm_ebc['D_REGION_ID'].astype(str).str[-3:])
    
    return(otm_ahc, otm_ebc)

def pivot_table_dataframe(dfa, output_folder):
    # Pivot table
    # Misspell in TOTAL_CLASS_WEGIHT
    today = datetime.date.today()
    table = pd.pivot_table(dfa, index=['SHIPMENT_GID', 'CITY_ORG', 
                                       'PROVINCE_CODE_ORG', 'POSTAL_CODE_ORG', 
                                       'COUNTRY_CODE3_GID_ORG', 'CITY_DES', 
                                       'PROVINCE_CODE_DES', 'POSTAL_CODE_DES', 
                                       'COUNTRY_CODE3_GID_DES', 'Stop Alt WT'], 
                           columns=['FAK Class Rank'], 
                           values=['FAK Class Structure for Rating', 'TOTAL_CLASS_WEGIHT'], 
                           aggfunc={'FAK Class Structure for Rating': np.mean, 
                                    'TOTAL_CLASS_WEGIHT': np.sum})
    
    table.columns.set_levels(['Average of FAK Class Structure for Rating (FAK Class Rank = 1)', 
                              'Average of FAK Class Structure for Rating (FAK Class Rank = 2)', 
                              'Sum of TOTAL_CLASS_WEIGHT (FAK CLass Rank = 1)', 
                              'Sum of TOTAL_CLASS_WEIGHT (FAK Class Rank = 2)'], level=1, inplace=True)
    table.index.names = ['SHIPMENT_GID', 'CITY_ORG', 
                                       'PROVINCE_CODE_ORG', 'POSTAL_CODE_ORG', 
                                       'COUNTRY_CODE3_GID_ORG', 'CITY_DES', 
                                       'PROVINCE_CODE_DES', 'POSTAL_CODE_DES', 
                                       'COUNTRY_CODE3_GID_DES', 'Stop Alt WT'] # Rename columns
    table.columns = table.columns.get_level_values(0)
    table.columns = ['Average of FAK Class Structure for Rating (FAK Class Rank = 1)', 
                     'Average of FAK Class Structure for Rating (FAK Class Rank = 2)',
                     'Sum of TOTAL_CLASS_WEIGHT (FAK Class Rank = 1)',
                     'Sum of TOTAL_CLASS_WEIGHT (FAK Class Rank = 2)']
    
    pivot_table_csv_path = (output_folder +'\\'+ 'SMC3 Pivot Table ' + str(today) +'.csv')
    
    # Exports
    table.to_csv(pivot_table_csv_path)

    smc3.ReSave(pivot_table_csv_path)
    
    return(pivot_table_csv_path)

def smc3_results(email, batchmarkpwd, user, pivot_table_csv_path):
    smc3_fpath = smc3.smc3_report(email, batchmarkpwd, user, pivot_table_csv_path)
    
    smc3_lookup = pd.read_csv(smc3_fpath)
    
    return(smc3_lookup)

def update_columns(dfa, inputs_folder,report):
    
    renaming_columns = open(inputs_folder + '\\' + "Rename_Columns.txt","r")
    for line in renaming_columns:
        line = line.strip().split(',')
        dfa = dfa.rename(columns={line[0]:line[1]})
    
    dfa['# Shipments Spot Rated'] = dfa['# Shipments Spot Rated'].replace(['N', 'Y'], [0, 1])
     
    if (report == 'TRUCKLOAD') | (report == 'TL'):
        
        dfa['Contract or Spot Rated'] = dfa['# Shipments Spot Rated'].replace([np.nan,0, 1], ['Undetermined','Contract', 'Spot Rated'])
        dfa['Planned SCAC = Actual SCAC']= np.where((dfa['Planned Service Provider SCAC'] == dfa['Service Provider SCAC']), 'Yes', 'No')
        dfa['AHC Flag'] = np.where(pd.notnull(dfa['AVG_HIST_CST_USD_AMT']), '', 'Not Available')
        dfa['EBC Flag'] = np.where(pd.notnull(dfa['EXP_BID_CST_USD_AMT']), '', 'Not Available')
        dfa['Questionable LHL Cost'] = dfa['Linehaul Cost'].apply(lambda x: '? LHL Cost' if x < 150 else 'Good')

        dfa = dfa.rename(columns={'EXP_BID_CST_USD_AMT':'EBC Cost'})
        dfa = dfa.rename(columns={'AVG_HIST_CST_USD_AMT':'AHC Cost'})
    
        del_columns = open(inputs_folder + '\\' + "LTL_remove_columns.txt","r") 
        for line in del_columns:
            line = line.strip('\n')
            dfa.drop([line], axis = 1, inplace=True) 
        
        #get fuel charge for missing value
        fuel_charge = dfa[['Fuel Surcharge Cost','Miles','Week Number']]
        fuel_charge = fuel_charge[(fuel_charge['Fuel Surcharge Cost'] != 0) &
                                  (fuel_charge['Miles'] != 0)]
        fuel_charge['rate'] = dfa['Fuel Surcharge Cost']/dfa['Miles']
        fuel_charge = fuel_charge[['rate','Week Number']]
        fuel_charge = fuel_charge.groupby(['Week Number']).mean()
        fuel_charge['Week Number'] = fuel_charge.index
        fuel_charge.index.names = ['Index']
        fuel_charge = fuel_charge.round(2)
        
        #merge values with original dataframe
        dfa = pd.merge(dfa, fuel_charge, on = 'Week Number',
                                   how='left')
        
        #mark for future purposes to change linehaul cost
        dfa['marker'] =np.where(dfa['Fuel Surcharge Cost']==0, 1,0)
        
        #if zero then do equation                 
        dfa['Fuel Surcharge Cost'] = np.where(dfa['Fuel Surcharge Cost']==0, 
                                           dfa['rate'] * dfa['Miles'], 
                                           dfa['Fuel Surcharge Cost'])
        
        #update linehaul cost where fuel surcharge were zeros
        dfa['Linehaul Cost'] = np.where(dfa['marker']==1, 
                                           dfa['Linehaul Cost'] - dfa['Fuel Surcharge Cost'], 
                                           dfa['Linehaul Cost'])
        
        miles_band = pd.read_excel(inputs_folder + '\\' + 'Jon_Miles_Band.xlsx')
        
        dfa['Origin_City_State'] = dfa['Origin City'] + " "+dfa['Origin State Province']
        
        miles_band['Origin_City_State'] = miles_band['Origin_City_State'].astype(str).str.upper()
 
        dfa = pd.merge_asof(dfa.sort_values('Miles'), 
                            miles_band.sort_values('mile'),
                            left_on = 'Miles',
                            right_on = 'mile',
                            by = 'Origin_City_State')
        
        dfa['Actual Cost mb'] = dfa['Miles'] * dfa['2019 Baseline RFP']
        
        dfa['Actual Cost mb'] = np.where(dfa['Actual Cost mb'] >=
                                        dfa['2019 Baseline Flat'], 
                                        dfa['Actual Cost mb'] , 
                                        dfa['2019 Baseline Flat'])
        
         #remove unused columns
        cols = ['rate','marker','Origin_City_State','2019 Baseline Flat','mile']
        dfa.drop(dfa[cols], axis = 1, inplace=True) 
        
        #update columns formatting
        #update deletion of columns
    
    return dfa
    
def Aggregating_Data_Part_Two(inputs_folder, output_folder, smc3_lookup, dfa, 
                              otm_ahc, otm_ebc):
    smc3_lookup.columns.values[:10] = ['Shipment Id', 'Origin City', 'Origin State Province', 'Origin Postal Code', 
                         'Origin Country', 'Destination City', 'Destination State Province', 'Destination Postal Code', 
                         'Destination Country', 'Stop Alt WT']
    
    dfa['Tariff'] = 'CZARLITE 9/14/2015'
    dfa['ST, Zip3 Lane'] = dfa['Origin State Province'] + ' ' + dfa['Origin Zip 3'] + ':' + dfa['Destination State Province'] + ' ' + dfa['Destination Zip 3']
    dfa['Shipment Count'] = 1 # If a Shipment Id has multiple rows for that one Shipment Id, then the other rows get 0 for Shipment Count
    
    i = 0 
    matches_counter = 0
    while i < len(dfa):
        sid = dfa.at[i, 'Shipment Id']
        # smc_index = smc3_lookup[smc3_lookup['Shipment Id'] == sid].index.item() # Returns index (row) number
        if len(smc3_lookup[smc3_lookup['Shipment Id'] == sid]) == 0: # If vlookup value doesn't exist
            print('No vlookup match for index', str(i) + ',', 'please review whether Shipment Ids are unique')
            matches_counter += 1
            print('Total number of non-matches thus far:', str(matches_counter))
        else:
            smc_index = smc3_lookup[smc3_lookup['Shipment Id'] == sid].index.item() # Returns index (row) number
            dfa.at[i, 'Tariff Base Cost (unformatted)'] = smc3_lookup.at[smc_index, '(1)TotalCharge']
            
        i += 1
    dfa.loc[dfa.duplicated(['Shipment Id']), 'Shipment Count'] = 0 
    dfa['Shipment Count'] = dfa['Shipment Count'].astype(int)
    
    # Rest of the columns are populated in a different while loop because their values to be filled are determined by existing data populated by the above loop
    i = 0
    matches_count = 0
    while i < len(dfa):
        state_zip_lane = dfa.at[i, 'ST, Zip3 Lane']
        if len(otm_ahc[otm_ahc['Lane'] == state_zip_lane]) == 0:
            print('Index', str(i))
            matches_count += 1
            print('Total # of non-matches:', str(matches_count))
            dfa.at[i, 'AHC Disc'] = 'None'
            dfa.at[i, 'AHC Min Chg'] = 'None'
            dfa.at[i, 'AHC Cost'] = 'None'
        else:  # Else if matches exist
            try:
                otm_ahc_index = otm_ahc[otm_ahc['Lane'] == state_zip_lane].index.item() # Returns index (row) number
            except ValueError:
                # print("The value of otm_ahc[otm_ahc['Lane'] == state_zip_lane].index is:" + str(otm_ahc[otm_ahc['Lane'] == state_zip_lane].index))
                print(otm_ahc[otm_ahc['Lane'] == state_zip_lane].index)
            # if pd.isna(otm_ahc.at[otm_ahc_index, 'CHARGE_MULTIPLIER_SCALAR']):
            #     dfa.at[i, 'AHC Disc (unformatted)'] = 'None'
            # else:
            dfa.at[i, 'AHC Disc (unformatted)'] = otm_ahc.at[otm_ahc_index, 'CHARGE_MULTIPLIER_SCALAR'] / 100
            # if pd.isna(otm_ahc.at[otm_ahc_index, 'RATE_GEO_MIN_COST']):
            #     dfa.at[i, 'AHC Min Chg (unformatted)'] = 'None'
            # else:
            dfa.at[i, 'AHC Min Chg (unformatted)'] = otm_ahc.at[otm_ahc_index, 'RATE_GEO_MIN_COST'] # float
            if pd.isna(dfa.at[i, 'Tariff Base Cost (unformatted)']) & pd.isna(dfa.at[i, 'AHC Disc (unformatted)']) & pd.isna(dfa.at[i, 'AHC Min Chg (unformatted)']):
                dfa.at[i, 'AHC Cost (unformatted)'] = 'None'
            else:
                compare_max = dfa.at[i, 'Tariff Base Cost (unformatted)'] * (1 - dfa.at[i, 'AHC Disc (unformatted)'])
                if compare_max >= dfa.at[i, 'AHC Min Chg (unformatted)']:
                    dfa.at[i, 'AHC Cost (unformatted)'] = compare_max
                else:
                    dfa.at[i, 'AHC Cost (unformatted)'] = dfa.at[i, 'AHC Min Chg (unformatted)']
        i += 1
    
    # dfa['AHC Disc (unformatted)'].replace(np.nan, 'None', inplace=True)
    # dfa['AHC Min Chg (unformatted)'].replace(np.nan, 'None', inplace=True)
    
    # EBC
    i = 0
    while i < len(dfa):
        state_zip_lane = dfa.at[i, 'ST, Zip3 Lane']
        if len(otm_ebc[otm_ebc['Lane'] == state_zip_lane]) == 0:
            dfa.at[i, 'EBC Disc'] = 'None'
            dfa.at[i, 'EBC Min Chg'] = 'None'
            dfa.at[i, 'EBC Cost'] = 'None'
        else:
            otm_ebc_index = otm_ebc[otm_ebc['Lane'] == state_zip_lane].index.item() # Returns index (row) number
            # if pd.isna(otm_ebc.at[otm_ebc_index, 'CHARGE_MULTIPLIER_SCALAR']):
            #     dfa.at[i, 'EBC Disc (unformatted)'] = 'None'
            # else:
            dfa.at[i, 'EBC Disc (unformatted)'] = otm_ebc.at[otm_ebc_index, 'CHARGE_MULTIPLIER_SCALAR'] / 100
            # if pd.isna(otm_ebc.at[otm_ebc_index, 'RATE_GEO_MIN_COST']):
            #     dfa.at[i, 'EBC Min Chg (unformatted)'] = 'None'
            # else:
            dfa.at[i, 'EBC Min Chg (unformatted)'] = otm_ebc.at[otm_ebc_index, 'RATE_GEO_MIN_COST'] # float
            if pd.isna(dfa.at[i, 'Tariff Base Cost (unformatted)']) & pd.isna(dfa.at[i, 'EBC Disc (unformatted)']) & pd.isna(dfa.at[i, 'EBC Min Chg (unformatted)']):
                dfa.at[i, 'EBC Cost (unformatted)'] = 'None'
            else:
                compare_max = dfa.at[i, 'Tariff Base Cost (unformatted)'] * (1 - dfa.at[i, 'EBC Disc (unformatted)'])
                if compare_max >= dfa.at[i, 'EBC Min Chg (unformatted)']:
                    dfa.at[i, 'EBC Cost (unformatted)'] = compare_max
                else:
                    dfa.at[i, 'EBC Cost (unformatted)'] = dfa.at[i, 'EBC Min Chg (unformatted)']
        i += 1
    
    # dfa['EBC Disc (unformatted)'].replace(np.nan, 'None', inplace=True)
    # dfa['EBC Min Chg (unformatted)'].replace(np.nan, 'None', inplace=True)    
    
    dfa['Contract or Spot Rated'] = dfa['# Shipments Spot Rated'].replace([np.nan,0, 1], ['Undetermined','Contract', 'Spot Rated'])
    dfa['Planned SCAC = Actual SCAC']= np.where((dfa['Planned Service Provider SCAC'] == dfa['Service Provider SCAC']), 'Yes', 'No')
    dfa['AHC Flag'] = np.where(pd.notnull(dfa['AHC Cost (unformatted)']), '', 'Not Available')
    dfa['EBC Flag'] = np.where(pd.notnull(dfa['EBC Cost (unformatted)']), '', 'Not Available')
    
    i = 0
    while i < len(dfa):
        # if pd.isna(dfa.at[i, 'AHC Cost (unformatted)']):
        #     dfa.at[i, 'AHC Flag'] = 'Not Available' # This is in a different loop since it depends on the data filled by the above loop
        # else:
        #     dfa.at[i, 'AHC Flag'] = ''
        # if pd.isna(dfa.at[i, 'EBC Cost (unformatted)']):
        #     dfa.at[i, 'EBC Flag'] = 'Not Available'
        # else:
        #     dfa.at[i, 'EBC Flag'] = ''
        
        # if dfa.at[i, '# Shipments Spot Rated'] == 0:
        #     dfa.at[i, 'Contract or Spot Rated'] = 'Contract'
        # elif dfa.at[i, '# Shipments Spot Rated'] == 1:
        #     dfa.at[i, 'Contract or Spot Rated'] = 'Spot Rated'
        # else:
        #     dfa.at[i, 'Contract or Spot Rated'] = ''
        
        # if dfa.at[i, 'Planned Service Provider SCAC'] == dfa.at[i, 'Service Provider SCAC']:
        #     dfa.at[i, 'Planned SCAC = Actual SCAC'] = 'Yes'
        # else:
        #     dfa.at[i, 'Planned SCAC = Actual SCAC'] = 'No'
        
        if (dfa.at[i, 'Linehaul Cost'] < 50) & (dfa.at[i, 'Tariff Base Cost (unformatted)'] < 1):
            dfa.at[i, 'Questionable LHL Cost or Tariff Base Cost'] = '? LHL and Tariff Base Cost'
        elif dfa.at[i, 'Linehaul Cost'] < 50:
            dfa.at[i, 'Questionable LHL Cost or Tariff Base Cost'] = '? LHL Cost'
        elif dfa.at[i, 'Tariff Base Cost (unformatted)'] < 1:
            dfa.at[i, 'Questionable LHL Cost or Tariff Base Cost'] = '? Tariff Base Cost'
        else:
            dfa.at[i, 'Questionable LHL Cost or Tariff Base Cost'] = 'Good'
        i += 1
    
    formatting_columns = open(inputs_folder + '\\' + "Formatting_Columns.txt","r")
    for line in formatting_columns:
        line = line.strip().split(',')

        if "Disc" in line[0]: # Disc is the only column with % values
            try:
                dfa[line[0]] = dfa[line[1]].apply(lambda x: "{:,.1%}".format(x))
            except ValueError: 
                print(line[0],line[1])
        else: # Others are columns with $ values
            dfa[line[0]] = '$' + dfa[line[1]].apply(lambda x: "{:.2f}".format(x))

            
    formatting_columns = open(inputs_folder + '\\' + "More_Formatting_Columns.txt","r")
    for each in formatting_columns: 
        each = each.strip('\n')
        if "Disc" in each: # Disc is the only column with % values
            dfa[each] = dfa[each].replace('nan%', 'None').replace('$nan%', 'None').replace('nan', 'None')
        else: # Others are columns with $ values
            dfa[each] = dfa[each].replace('$nan', 'None').replace('nan', 'None')
            
    del_columns = open(inputs_folder + '\\' + "Unwanted_Columns_Unformatted.txt","r") 
    for line in del_columns:
        line = line.strip('\n')
        dfa.drop([line], axis = 1, inplace=True) 
    
    # Reorder columns
    move_to_end = ['AHC Flag', 'EBC Flag', 'Contract or Spot Rated', 
                                     'Planned SCAC = Actual SCAC', 'Questionable LHL Cost or Tariff Base Cost']
    dfa = dfa[[c for c in dfa if c not in move_to_end]
        + move_to_end]
    
    return dfa

def create_reports(output_folder,dfa,report):
    today = datetime.date.today()
    
    if report == "LTL":
        report_filepath = (output_folder + '\\' + 'Report ' + str(today) + '.xlsx')
        
        dfa.to_excel(report_filepath, index=False)
        
    elif (report == 'TRUCKLOAD') | (report == 'TL'):
        report_filepath = (output_folder + '\\' + 'Report_TL ' + str(today) + '.xlsx')
        
        dfa.to_excel(report_filepath, index=False)

def main():
    start_time = time.time()
    
    if (report == 'TRUCKLOAD') | (report == 'TL'):
        
        ecdw, hstp_new, audit_rate = Oracle_Queries(user, report, sql_folder)
        
        dfa = Aggregating_Data_Part_One(ecdw, hstp_new, inputs_folder)
        
        dfa = update_columns(dfa, inputs_folder,report)
        
        create_reports(output_folder,dfa,report)

    elif report == 'LTL':
        
        ecdw, hstp_new, audit_rate = Oracle_Queries(user, report, sql_folder)
        
        dfa = Aggregating_Data_Part_One(ecdw, hstp_new, inputs_folder)
        
        dfa = nmfc(dfa, inputs_folder)
        
        otm_ahc, otm_ebc = audit_rate_dataframes(audit_rate)
        
        pivot_table_csv_path = pivot_table_dataframe(dfa, output_folder)
        
        smc3_lookup = smc3_results(email, batchmarkpwd, user, pivot_table_csv_path)
        
        dfa = update_columns(dfa, inputs_folder,report)
        
        dfa = Aggregating_Data_Part_Two(inputs_folder, output_folder, smc3_lookup, dfa,
                                  otm_ahc, otm_ebc)
        
        create_reports(output_folder,dfa,report)
 
    print('\nTotal runtime, including time for user input, took %s minute(s) and %d second(s)' 
      % (math.floor((time.time() - start_time) / 60), (time.time() - start_time) % 60), '\n')

if __name__ == "__main__":
    main()
