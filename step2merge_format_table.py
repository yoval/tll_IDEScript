# -*- coding: utf-8 -*-
"""
Created on Tue May 21 17:30:18 2024

@author: Administrator
"""
from my_module import list_files
import pandas as pd
import configparser


# 读取配置文件
config = configparser.ConfigParser()
config.read('config.ini', encoding='utf-8')

# 获取配置值
folder = config.get('xiaoshou', 'folder')
syj = config.get('xiaoshou', 'shouyinji')
fileList = list_files(folder)


filename_patterns = ['本期_格式化', '环比期_格式化', '同比期_格式化']
for pattern in filename_patterns:  
    files = [file for file in fileList if file.startswith(pattern)]
    if len(files) == 0:
        continue
    elif len(files) == 2:
        df1 = pd.read_excel(files[0])
        df2 = pd.read_excel(files[1])
        merged_df = pd.merge(df1, df2, on='门店编码', suffixes=('_1', '_2'))
        columns_to_add = [col for col in merged_df.columns if col not in ['门店编码']]
        for col in columns_to_add:
            col1, col2 = col.split('_')[0] + '_1', col.split('_')[0] + '_2'
            merged_df[col.split('_')[0]] = merged_df[col1] + merged_df[col2]
        merged_df = merged_df.drop([col for col in merged_df.columns if '_1' in col or '_2' in col], axis=1)
        merged_df.to_excel(f'{folder}\\结果_{pattern}合计收银.xlsx',index=False)
    elif len(files) == 1:
        df1 = pd.read_excel(files[0])
        df1.to_excel(f'{folder}\\{pattern}合计收银.xlsx',index=False)
    else:
        print(len(files))
        print('文件数量有误！')