# -*- coding: utf-8 -*-
"""
Created on Tue May 21 15:15:11 2024

@author: Administrator

加盟店汇总表制作
美团：综合营业统计
哗啦啦：142
"""
from my_module import format_meituan_table,format_zhongtai_table,format_hualala_table,list_files
import pandas as pd
import configparser
import re

# 读取配置文件
config = configparser.ConfigParser()
config.read('config.ini', encoding='utf-8')

# 获取配置值
folder = config.get('xiaoshou', 'folder')
syj = config.get('xiaoshou', 'shouyinji')
fileList = list_files(folder)

columns_to_read = ['收银机ID', '门店编码']
syj_df = pd.read_excel(syj, usecols=columns_to_read)

# 批量格式化
meituan_files = [file for file in fileList if '安徽汇旺餐饮管理有限公司_综合营业统计' in file]
for meituan_file in meituan_files:
    df_check = pd.read_excel(meituan_file, nrows=2)
    pattern = r'\d{4}/\d{2}/\d{2}'
    dates = re.findall(pattern,  df_check.iloc[0][0])
    shiduan = dates[0] + '~' + dates[1]
    shiduan = shiduan.replace('/', '')
    df = format_meituan_table(meituan_file)
    df.to_excel(f'{folder}\\格式化_美团收银_{shiduan}.xlsx',index=False)

hualala_files = [file for file in fileList if '142渠道销售统计表' in file]
for hualala_file in hualala_files:
    df_check = pd.read_excel(hualala_file, nrows=2)
    df = format_hualala_table(hualala_file)
    pattern = r'(\d{4})(\d{2})(\d{2})--(\d{4})(\d{2})(\d{2})'
    date_range = re.search(pattern, df_check.columns[0])
    start_date = f"{date_range.group(1)}/{date_range.group(2)}/{date_range.group(3)}"
    end_date = f"{date_range.group(4)}/{date_range.group(5)}/{date_range.group(6)}"
    shiduan = f"{start_date}~{end_date}"
    shiduan = shiduan.replace('/', '')
    df = format_hualala_table(hualala_file)
    df.to_excel(f'{folder}\\格式化_哗啦啦收银_{shiduan}.xlsx',index=False)

zhongtai_files = [file for file in fileList if file.endswith('.csv')]
for zhongtai_file in zhongtai_files:
    df_check = pd.read_csv(zhongtai_file, nrows=2,encoding='gbk')
    shiduan = df_check['时段'][0]
    shiduan = shiduan.replace('\t','')
    df = format_zhongtai_table(zhongtai_file)
    df.to_excel(f'{folder}\\格式化_中台收银_{shiduan}.xlsx',index=False)