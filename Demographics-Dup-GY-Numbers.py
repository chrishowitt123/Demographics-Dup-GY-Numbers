
import pandas as pd
import os
import itertools
import xlsxwriter
from rapidfuzz import fuzz
from datetime import date
from ordered_set import OrderedSet

os.chdir('M:\GY number duplicates')
pd.set_option('display.max_columns', 1000, 'display.width', 1000, 'display.max_rows',1000)
demogs = df = pd.read_csv('GY_update_25.04.22.txt', '\t')
# Only rows with GY numbers
demogs = demogs[demogs['GY_Number'].str.contains('GY')]

# Ensure no duplicates
demogs = demogs.drop_duplicates() 

# Remove NaNs
demogs.fillna('', inplace=True)

# Create DatFrame for fuxxy matching (pairs) and DataFrame with more than 2 instances 
demogs_pairs = demogs[demogs.groupby("GY_Number")['GY_Number'].transform('size') == 2]
demogs_more_than_2_instances = demogs[demogs.groupby("GY_Number")['GY_Number'].transform('size') > 2]

# Group paird DataFrame by SSN
groups_paris = demogs_pairs.groupby('GY_Number') 

master_list = []
for group_name, df_group in groups_paris:
    # Each group as a list (group) of lists (rows)
    gy_group_list = df_group.values.tolist() 
    master_list.append(gy_group_list)

# Define new columns
cols = demogs.columns.tolist()
cols.append('Ratio')
cols.append('Diffs')
cols.append('NoInstances')

# Create new DataFrame to hold results
df_results = pd.DataFrame(columns = cols)

master_score = []

# For each list (group) in master_list
for i in range(len(master_list)):
    score_list = []
    avg_score = []
    diffA_list = []
    diffB_list = []
    
    # For each row in group minus URN
    for x, y in zip(master_list[i][0][1:], master_list[i][1][1:]): 
        score = fuzz.ratio(x, y)
        score_list.append(score)
        
        # Get diffrences between fields for specific field using OrderedSet for clarity
        splitB = OrderedSet(x.split(' '))
        splitA = OrderedSet(y.split(' '))
        diffA = splitA - splitB
        diffB = splitB - splitA
        if diffA != '':
            diffA = ' '.join(diffA) 
            diffA_list.append(diffA)
        if diffB != '':
            diffB = ' '.join(diffB) 
            diffB_list.append(diffB)
    
    # Score out of 100
    avg_score = sum(score_list)/len(score_list)
    
    # Append new information to orignal rows
    master_list[i][0].append(int(avg_score))
    master_list[i][1].append(int(avg_score))

    # Remove empty strings
    diffA_list = list(filter(None, diffA_list))
    diffB_list = list(filter(None, diffB_list))

    diffA_string = '        '.join(diffA_list)
    diffB_string = '        '.join(diffB_list)
    
    master_list[i][0].append(diffB_string)
    master_list[i][1].append(diffA_string)
    
    master_list[i][0].append('')
    master_list[i][1].append('')

    #Append final lists to DataFrame
    df_results.loc[len(df_results)] = master_list[i][0]
    df_results.loc[len(df_results)] = master_list[i][1]
    
# Append more thank 2 instances DataFrameto results
df_results = df_results.append(demogs_more_than_2_instances)

# Count number of instances of GY numbers and add to DataFrame
df_results['NoInstances'] = df_results['GY_Number'].map(df_results['GY_Number'].value_counts())

# Sort results
df_results.sort_values(by= ['Ratio', 'GY_Number'], ascending = [False, False], inplace = True) 

#Add commments column
df_results['Comments'] = ''

# Re-order Columns
df_results = df_results[[
 'URN',
 'FirstName',
 'LastName',
 'DOB',
 'Address1',
 'Address2',
 'Gender',
 'GY_Number',
 'Ratio',
 'Diffs',
 'NoInstances',
 'Comments']]

# To Excel..
# Todays date
today = date.today()
date = today.strftime("%d.%m.%Y")

# Create a Pandas Excel writer using XlsxWriter as the engine
writer = pd.ExcelWriter(f'GY_Number_Dups_{date}.xlsx', engine='xlsxwriter') 

# Convert  dataframe to  XlsxWriter Excel object. Turn off the default
# header and index and skip one row to allow us to insert a user defined
# header
df_results.to_excel(writer, sheet_name= f'GY_Number_Dups_{date}', startrow=1, header=False, index=False)

# Get the xlsxwriter workbook and worksheet objects
worksheet = writer.sheets[f'GY_Number_Dups_{date}']

# Get  dimensions of df_results
(max_row, max_col) = df_results.shape 

# Create a list of column headers, to use in add_table() 
column_settings = [] 
for header in df_results.columns:
    column_settings.append({'header': header})

# Add table to Excel as per dimensions of results_df
worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings, 'style': 'Table Style Light 11'})

# Make the columns wider for clarity
worksheet.set_column(0, max_col - 1, 12) 

# Close the Pandas Excel writer and output the Excel file
writer.save() 

print('Finished!')
Finished!
 
