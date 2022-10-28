# -*- coding: utf-8 -*-
"""
Created on Wed Oct 26 18:08:00 2022

@author: jacob
"""
import time
from AutoReport.Replace_word import *



print(f"開始執行程式")
CallVBA()
print(f'緩衝兩秒')
time.sleep(2)
print(f'開始製作稽核報告Word檔')
Replace_Word()
