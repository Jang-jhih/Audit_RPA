import pandas as pd
import re
import numpy as np
from pathlib import Path
import os

def CreatTXT(dirpath):

    # allroot =[]
    # for root, dirs, files in os.walk(r'D:\OneDrive - 全支付電子支付股份有限公司\1.建立調閱路徑\FD'):
    #     allroot.append(root)
    
    txtcontent = [_.split('\\')[-1] for _ in dirpath]
    
    dirs = [_.split('\\') for _ in dirpath]
    dirs = pd.DataFrame(dirs).dropna()
    
    
    
    dirs['txt'] = '檔案說明.txt'
    dirs.reset_index(drop = True,inplace = True)
    dirs = dirs.T
    dirs.reset_index(drop = True,inplace = True)
    
    
    Description=[]
    for column in list(dirs.columns):
        Description.append('\\'.join(list(dirs[column])))
        
        
    for txt,content in zip(Description,txtcontent):
        # print(txt)
        
        path = Path(txt)
        # path.exists()
        path.touch()
        

        f = open(txt, 'w')
        f.write(f'{content}\n\n檔案說明：')
        f.close()
        # path.exists()