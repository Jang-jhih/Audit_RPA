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





def ClearNullDir(file, test = False):

    if test == True:
        maindir = os.path.join('test')
    else:
        maindir = r'\\192.168.248.10\IA_Evidence'
    


def referance(Department,file,item):
    return Department+file.split('.')[0][-3:]+file.split('.')[1].zfill(2)+'-'+item
    
    #%%
    # 測試
    # test
    # 是否建立檔案說明
    # creattxt
    # 查單位代號
    # Department
def Auto_mkdir(file, test = False,creattxt = False):
    #用檔案的中文單位搜尋英文代號
    file = os.path.join('document',file)
    Department = pd.read_excel(os.path.join('condition', 'Department.xlsx'), sheet_name='單位代號')
    Department_Fliter = file.split('_')[1].split('.')[0]
    Department = Department[Department['單位'] == Department_Fliter]['英文代號'].tolist()[0]

    #%%
    # 調閱清單路徑
    df = pd.read_excel(file, skiprows= 2).astype('str')
    
    # 標註用號碼
    df['號碼'] = df['項次']
    
    #%% 清除數字
    df['項次'] = [re.sub('[0-9]','',_) for _ in list(df['項次'])]
    
    # 把文字'nan'轉換成空值
    df = df.apply(lambda x:x.replace('nan',np.nan))
    # 把空白轉成空
    df = df.replace(r'^\s*$', np.nan, regex=True)
    
    #%%專案查核用
    if len(df[~df['項次'].isnull()]) <=1:
        df.fillna(method='ffill', inplace = True)
        df = df[~df[df.columns[1]].isnull()]
        df['底稿索引'] = referance(Department=Department,file=file,item=df['號碼'])
        
        ItemList =(df['項次'] +'\\'+ df['底稿索引']).tolist()
    
    #%%例行查核用
    else:
        # 填滿空值
        df.fillna(method='ffill', inplace = True)
        
        
        #調整資料夾名稱
        
        df = df[~df[df.columns[1]].isnull()]
        convert = pd.read_excel(os.path.join('condition', 'Department.xlsx'), sheet_name='數字轉換')
        for old,new in zip(convert[convert.columns[1]],convert[convert.columns[0]]):
            df['項次']=df['項次'].str.replace(old,str(new))

        df['項次'] = df['項次'].apply(lambda x:referance(Department, file, item=x.split('、')[0]))
        
        

        
        df.drop_duplicates(subset = '查核標的及程序', keep = 'last', inplace = True) 
        # 異常的號碼拿掉
        df = df[df['號碼'].apply(lambda x:len(x)<3)]
        
        
        
        # df['查核標的及程序'] = df['號碼'] +'、'+ df['查核標的及程序']
        df['查核標的及程序'] = df['項次'] +'-'+df['號碼']
        
        
        df.set_index('項次' ,inplace = True)
        df = df[['查核標的及程序']]
        

        df['查核標的及程序'] = [_.replace('?','？').replace('承上，','').split('\n')[0] for _ in list(df['查核標的及程序'])]
        df.reset_index(inplace = True)
        df.insert(0, "母資料夾", df['項次'].apply(lambda x:x[:7]))
        
        
        df = df.T


        ItemList = []
        for name,f in df.items():
            ItemList.append('\\'.join(list(df[name])))
            

    #%%
    if test == True:
        maindir = os.path.join('test')
        
    else:
        maindir = r'\\192.168.248.10\IA_Evidence'
    
    
    dirpath = []
    for dir in ItemList:
        tmp = f'{maindir}\{Department}\{dir}'
        dirpath.append(tmp)
        path = Path(tmp)
        print(dir)
        try:
            path.mkdir(parents=True, exist_ok=True)
        except:
            path = Path(tmp.split('，')[0])
            path.mkdir(parents=True, exist_ok=True)
    #%%建立TXT
    
    if creattxt==True:
        CreatTXT(dirpath)



import os

def del_none_folder(path):
    max_len = len(path.split('\\'))
    for folder,subfolder,file in os.walk(path):
        # 获取最大路径长度
        if len(folder.split('\\')) > max_len:
            max_len = len(folder.split('\\'))
    # 获取相对路径的长度
    max_len = max_len - len(path.split('\\'))
    # 对长度进行遍历删除
    for i in range(max_len):
        for fo,sf,f in os.walk(path):
            # 如果目录下没有文件夹和文件，说明此路径是空的，就删掉
            if not sf and not f:
                os.rmdir(fo)
                print('Deleted: {}'.format(fo))
    print('Done.')

