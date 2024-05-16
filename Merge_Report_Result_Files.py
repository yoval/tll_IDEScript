# -*- coding: utf-8 -*-
"""
Created on Thu May 16 13:55:09 2024

@author: Administrator
报货脚本 新品报货自动处理、单品报货自动处理(90天) 处理结果合并
"""
import pandas as pd
import glob,os,time
import configparser

config = configparser.ConfigParser()
config.read('config.ini', encoding='utf-8')
outputs_folder = config.get('WebScript', 'outputs_folder')

# 读取底表sheet
def Read_bottom_tables(file):
    print(file)
    excel_file = pd.ExcelFile(file)
    sheet_names = excel_file.sheet_names
    bottom_tables = [sheet for sheet in sheet_names if '底表' in sheet][0]
    df = pd.read_excel(file, sheet_name = bottom_tables)
    food_name = bottom_tables.replace('底表','')
    return food_name, df
now = time.strftime('%Y%m%d_%H%M', time.localtime())
today_ = time.strftime('%Y%m%d_', time.localtime())
fileList = glob.glob(os.path.join(outputs_folder, f'*报货信息_{today_}*.xlsx')) 
file = fileList[0]
food_name,df1_original = Read_bottom_tables(file)
df1 = df1_original.loc[: , ['门店编码', '门店名称', '大区经理', '省区经理', '区域经理', '运营状态','省', f'{food_name}报货周期']]

for file in fileList[1:]:
    food_name_tb,df_original = Read_bottom_tables(file)
    food_name = food_name + '&' + food_name_tb
    df = df_original.loc[:,['门店编码',f'{food_name_tb}报货周期']]
    df1= pd.merge(df1, df, how='left', on='门店编码')
    
df1.to_excel(f'{outputs_folder}\\{food_name}报货汇总_{now}.xlsx',index=False)