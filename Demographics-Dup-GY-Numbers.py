import pandas as pd
import os
import itertools
import xlsxwriter
from rapidfuzz import fuzz
from datetime import date
from ordered_set import OrderedSet
pd.set_option('display.max_columns', 1000, 'display.width', 1000, 'display.max_rows',1000)

# set CWD
os.chdir('M:\GY number duplicates')

# data to DataFrame
demogs = df = pd.read_csv('GY_update_25.04.22.txt', '\t')

# only rows with GY numbers
demogs = demogs[demogs['GY_Number'].str.contains('GY')]

# ensure no duplicates
demogs = demogs.drop_duplicates() 

# remove NaNs
demogs.fillna('', inplace=True)

# create DatFrame for fuzzy matching (pairs) and DataFrame with more than 2 instances 
demogs_pairs = demogs[demogs.groupby("GY_Number")['GY_Number'].transform('size') == 2]
demogs_more_than_2_instances = demogs[demogs.groupby("GY_Number")['GY_Number'].transform('size') > 2]

# group paired DataFrame by GY number
groups_paris = demogs_pairs.groupby('GY_Number') 

master_list = []
for group_name, df_group in groups_paris:
    # each group as a list (group) of lists (rows)
    gy_group_list = df_group.values.tolist() 
    master_list.append(gy_group_list)

# define new columns
cols = demogs.columns.tolist()
cols.append('Ratio')
cols.append('Diffs')
cols.append('NoInstances')

# create new DataFrame to hold results
df_results = pd.DataFrame(columns = cols)

master_score = []
# for each list (group) in master_list
for i in range(len(master_list)):
    score_list = []
    avg_score = []
    diffA_list = []
    diffB_list = []
    
    # for each row in group minus URN
    for x, y in zip(master_list[i][0][1:], master_list[i][1][1:]): 
        score = fuzz.ratio(x, y)
        score_list.append(score)
        
        # get differences between specific field using OrderedSet for clarity
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
    
    # score out of 100
    avg_score = sum(score_list)/len(score_list)
    
    # append new information to original rows
    master_list[i][0].append(int(avg_score))
    master_list[i][1].append(int(avg_score))

    # remove empty strings
    diffA_list = list(filter(None, diffA_list))
    diffB_list = list(filter(None, diffB_list))

    # join fields using enough space for clarity
    diffA_string = '        '.join(diffA_list)
    diffB_string = '        '.join(diffB_list)
    
    # append to master list
    master_list[i][0].append(diffB_string)
    master_list[i][1].append(diffA_string)
    
    master_list[i][0].append('')
    master_list[i][1].append('')

    #append final lists to DataFrame
    df_results.loc[len(df_results)] = master_list[i][0]
    df_results.loc[len(df_results)] = master_list[i][1]
    
# append more than 2 instances DataFrameto results
df_results = df_results.append(demogs_more_than_2_instances)

# count number of instances of GY numbers and add to DataFrame
df_results['NoInstances'] = df_results['GY_Number'].map(df_results['GY_Number'].value_counts())

# sort results
df_results.sort_values(by= ['Ratio', 'GY_Number'], ascending = [False, False], inplace = True) 

# add comments column
df_results['Comments'] = ''

# re-order Columns
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

# to Excel
# todays date
today = date.today()
date = today.strftime("%d.%m.%Y")

# create a Pandas Excel writer using XlsxWriter as the engine
writer = pd.ExcelWriter(f'GY_Number_Dups_{date}.xlsx', engine='xlsxwriter') 

# convert dataframe to XlsxWriter Excel object. Turn off the default
# header and index and skip one row to allow the insertion of defined header
df_results.to_excel(writer, sheet_name= f'GY_Number_Dups_{date}', startrow=1, header=False, index=False)

# get the xlsxwriter workbook and worksheet objects
worksheet = writer.sheets[f'GY_Number_Dups_{date}']

# get dimensions of df_results
(max_row, max_col) = df_results.shape 

# create a list of column headers, to use after table added 
column_settings = [] 
for header in df_results.columns:
    column_settings.append({'header': header})

# add table to Excel as per dimensions of results_df
worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings, 'style': 'Table Style Light 11'})

# make the columns wider for clarity
worksheet.set_column(0, max_col - 1, 12) 

# close the Pandas Excel writer and output the Excel file
writer.save() 

print('Finished!')
 
