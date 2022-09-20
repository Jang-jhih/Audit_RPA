# -*- coding: utf-8 -*-
"""
Created on Thu Jul 14 11:01:14 2022

@author: jacob
"""
from audit_tool.law_search import *
import os

# 建立資料夾
if not os.path.isdir(os.path.join('datasource')):
    os.mkdir(os.path.join('datasource'))  
    
# 存放原始PDF
if not os.path.isdir(os.path.join('datasource','raw')):
    os.mkdir(os.path.join('datasource','raw'))
    
    
# 存放轉換後檔案
if not os.path.isdir(os.path.join('datasource','base')):
    os.mkdir(os.path.join('datasource','base'))      
  
if not os.path.isdir(os.path.join('search')):
    os.mkdir(os.path.join('search')) 
    
    
#%%
Related_legislation = ['電子支付','金融機構防制洗錢辦法']

for key_word in Related_legislation:
    crawl_law(key_word)

#關鍵字搜尋

key_words = pd.read_excel('自選關鍵字.xlsx')['關鍵字'].tolist()
key_words = [i for i in key_words if str(i)!='nan']
[search_keyword(key_word) for key_word in key_words ]





#%%

import glob
import os
import pandas as pd
from datetime import date

all_file = os.listdir(os.path.join('search'))
all_file = [os.path.join('search',_) for _ in all_file]

df = pd.DataFrame()
for i in all_file:
    new_df = pd.read_csv(i)
    new_df['path'] = i
    df = pd.concat([new_df,df])

df = df[['path','條號','內容']]



today = str(date.today())
filename = os.path.join('search',f"{today}.xlsx")

df.to_excel(filename ,index = False)
