import pandas as pd
import os
import xlsxwriter

from config import path_config

base_path = path_config.path_config()
old_data_path = base_path.origin_cate_path
old_data_files = [file for file in os.listdir(old_data_path)]

# new_age = pd.read_excel(r'C:\Users\PL_Dell3668_One\Desktop\work_projects\update_zipnomic_data\mapping\new_merged.xlsx')

# load old data and figure out the frame of the excel

Age_Sex_old = pd.ExcelFile(old_data_path+'\\'+old_data_files[0])
#
#
# # print(HH_Income.sheet_names)
#
Age_Sex_old_df = Age_Sex_old.parse('2015')
# print(Age_Sex_old_df.head())
#
old_indx = Age_Sex_old_df.index
print('old_index is: ',old_indx)
old_cols = Age_Sex_old_df.columns
print('old_cols is: ',old_cols)

print(Age_Sex_old_df.iloc[0])