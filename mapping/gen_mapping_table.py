import os, xlsxwriter, time
import pandas as pd
from openpyxl import load_workbook
import numpy as np

from config import path_config


base_path = path_config.path_config()

cat_data_16_path = base_path.cate_path
meta_data_16_path = base_path.cate_meta_pth
old_data_path = base_path.origin_cate_path
mapping_path = base_path.mapping_table_path

cat_files = [file for file in os.listdir(cat_data_16_path)]
meta_files = [file for file in os.listdir(meta_data_16_path)]
old_data_files = [file for file in os.listdir(old_data_path)]

# process one category file -- Age_Sex



# read csv file into DataFrame
mapping_label_df = pd.read_excel(mapping_path+'\\'+'Mapping Table With Source.xlsx')
mapping_zip_df = pd.read_excel(mapping_path+'\\'+'state_zip_16.xlsx', converters={'ZCTA5': lambda x: str(x)})

mapping_label_df = mapping_label_df[mapping_label_df['Level 1'] == 'Age and Sex']

Var_Geo_df = mapping_label_df[['Variable Code', 'GEO.display-label']]
# print(Var_Geo_df.columns)

Age_Sex_16_df = pd.read_csv(cat_data_16_path+'\\'+cat_files[0])

# access the geo_labels and correspond to geo_display
cols = Age_Sex_16_df.columns
geo_label = cols[3:]
geo_display = Age_Sex_16_df.iloc[0][3:]

Age_Sex_16_df = Age_Sex_16_df.iloc[:, 2:]
Age_Sex_16_df = Age_Sex_16_df.T
# print(Age_Sex_16_df.head())
Age_Sex_16_df.columns = Age_Sex_16_df.iloc[0]
# Age_Sex_16_df.reindex(Age_Sex_16_df.index.drop(0))
Age_Sex_16_df = Age_Sex_16_df.iloc[1:]
Age_Sex_16_df['GEO.display-label'] = geo_label

# print(Age_Sex_16_df['GEO.display-label'][0])

# merge two tables
result = pd.merge(Var_Geo_df, Age_Sex_16_df,  how='inner', on='GEO.display-label')
result = result.T
result = result.drop('GEO.display-label')
zipcode_lst = result.index.tolist()[2:]
zipcode_lst = [zipcode.split(' ') for zipcode in zipcode_lst]
zipcode_lst = [zipcode_lst[i][1] for i in range(len(zipcode_lst))]

result.reset_index(inplace=True)
result['index'][0] = 'ZIP'
result['index'][1] = 'ZIP'
result['index'][2:] = zipcode_lst
result.columns = result.iloc[0]
result = result.iloc[1:]
# result.set_index(None, inplace=True)

# print(res_cols)
# print(result.tail())
# print(result.head())


# load old data and figure out the frame of the excel

# Age_Sex_old = pd.ExcelFile(old_data_path+'\\'+old_data_files[0])
# #
# #
# # # print(HH_Income.sheet_names)
# #
# Age_Sex_old_df = Age_Sex_old.parse('2015')
# # print(Age_Sex_old_df.head())
# #
# old_indx = Age_Sex_old_df.index
# print('old_index is: ',old_indx)
# old_cols = Age_Sex_old_df.columns
# print('old_cols is: ',old_cols)

# process zip state mapping

mapping_zip_df = mapping_zip_df.rename(columns={'STUSAB':'State', 'ZCTA5':'ZIP'})
mapping_zip_df = mapping_zip_df[['State', 'ZIP']]
# mapping_zip_df['ZIP'] = mapping_zip_df['ZIP'].astype(str)
# print(mapping_zip_df.head())

# new_result = pd.merge(mapping_zip_df, result, how='inner', on='ZIP')

# print(result.head())
# result.columns = result.iloc[2]
# result = result.ilco[3:]
# print(result.head())
merged = pd.merge(mapping_zip_df, result, on='ZIP', how='right')
merged = merged.iloc[::-1]
var_des = merged.iloc[0]
new_merged = merged.drop([0])
new_merged.loc[-1] = var_des
new_merged.index = new_merged.index+1
new_merged.sort_index(inplace=True)
new_merged = new_merged.drop([len(new_merged)])

#####################################################################

# new_merged['State'][0] = 'State'
# test = new_merged.rename(columns={'State': None, 'ZIP': None})
# test = test.reset_index()
# test = test.drop(['index'], axis=1)
#
# new_test = test.set_index([test.iloc[:,0], test.iloc[:,1]])
#
# a = new_test.drop([None], axis=1)
# writer = pd.ExcelWriter('new_merged_2.xlsx', engine='xlsxwriter')
# a.to_excel(writer, sheet_name='2016')
# writer.save()

##########################################################################
# add 2016 data sheet to exist excel.

writer = pd.ExcelWriter(old_data_path+'\\'+old_data_files[0], engine='openpyxl')

try:

    # try to open an existing workbook
    writer.book = load_workbook(old_data_path+'\\'+old_data_files[0])
# copy existing sheets
    writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
    print(writer.sheets)
except IOError:
    # file does not exist yet, we will create it
    pass

# write out the new sheet
result.to_excel(writer, sheet_name='2016')

# save the workbook
writer.save()


