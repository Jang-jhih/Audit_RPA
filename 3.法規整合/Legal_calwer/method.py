import os
import pandas as pd
import pdfplumber
from Legal_calwer.Calwer_law import *
import warnings
warnings.filterwarnings("ignore")

def ConcatAllLegal():
    #%%這裡是針對PDF的法規
    df_normal = Newline('一般')
    df_appendix = Newline('金管會附件')
    
    df = pd.concat([df_normal,df_appendix])
    df['內容']=df['內容'].apply(lambda x:x.replace('。','。\n').replace('：','：\n'))
    add_to_csv(df)
    
    
    df = pd.read_csv(os.path.join('datasource','base.csv'))
    

    
    df.to_csv(os.path.join('datasource','final.csv') ,encoding = 'utf-8-sig' ,index=False)


    #%% 這裡是做法規爬蟲
    Related_legislation = ['電子支付','金融機構防制洗錢辦法']
    
    for key_word in Related_legislation:
        crawl_law(key_word)


def Newline(method):

    
    ConvertNumber = pd.read_excel(os.path.join('convert','convert_newline.xlsx') ,sheet_name='number',dtype={'old':str,'new':str}).astype('str')



    legallist = os.listdir( os.path.join('legal',method))
    content = pd.DataFrame()
    for path in legallist:
        print(path)

        File = open(os.path.join('legal',method,path),'rb')
        

        text = []
        with pdfplumber.open(File) as pdf:
            for page in pdf.pages:
                text.append(page.extract_text())

        text = ''.join(text)
        


        if method == '金管會附件':
            text = text.split('\n \n')
        else:
            text = text.replace('\n','nn').replace('nn第 ','@@第').replace(' ','nn')
            for old,new in zip(list(ConvertNumber['old']),list(ConvertNumber['new'])):
                text = text.replace(old,new).replace('nan','')

            text = text.replace('n','').replace('!!','\n').split('@@')





        text = pd.DataFrame(text)
        
        text['法規名稱'] =  path
        
        text.columns = ['內容', '法規名稱']
        text = text[['法規名稱', '內容']]
    
        content = pd.concat([content,text])

        
    return content


def SearchWord(department):
    
    old_df = pd.read_csv(os.path.join('datasource','final.csv'))
    key_words = list(pd.read_excel('review.xlsx', sheet_name='自選關鍵字')[department])
    key_words = [i for i in key_words if str(i)!='nan']
    
    
    AllKeyWord = pd.DataFrame()
    for key_word in key_words:
        print(key_word)
        df = old_df[old_df['內容'].str.contains(key_word)]
        df['關鍵字'] = key_word
        AllKeyWord = pd.concat([AllKeyWord,df])
    
    
    ExcludeWords = list(pd.read_excel('review.xlsx', sheet_name='排除法規'))
    
    finalLegal = pd.DataFrame()
    for ExcludeWord in ExcludeWords:
        AllKeyWord = AllKeyWord[~AllKeyWord['法規名稱'].str.contains(ExcludeWord)]
        finalLegal = pd.concat([finalLegal,AllKeyWord])
        
    finalLegal = finalLegal[['關鍵字','法規名稱', '內容']]
    finalLegal.to_csv(os.path.join('datasource','search.csv'),index = False, encoding='utf-8-sig')
