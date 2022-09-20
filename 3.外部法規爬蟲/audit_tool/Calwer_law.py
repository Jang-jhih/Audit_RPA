import os
import requests
from bs4 import BeautifulSoup
import re
from fake_useragent import UserAgent
import ssl
import pandas as pd
import urllib


   

def get_requests(url):
    ua = UserAgent()
    user_agent = ua.random  #隨機更新agent
    headers = {'user-agent' : user_agent}
    context = ssl._create_unverified_context() #取得SSL   
    res = urllib.request.Request(url,headers=headers) # 發送請求
    res = urllib.request.urlopen(res,context=context).read() #讀取Http
    res = res.decode('utf-8') #調整編碼    
    return res


def crawl_law(key_word):
    url = 'https://law.moj.gov.tw/Law/LawSearchResult.aspx?ty=ONEBAR&kw='+key_word
    content =  requests.get(url).text
    soup = BeautifulSoup(content,'html.parser')
    
    
    all_link = []
    for link in soup.find_all('a', href=True,id="hlkLawLink"):
        #這裡只有抓每個法規的編號
       all_link.append('https://law.moj.gov.tw/LawClass/LawAll.aspx?pcode='+re.search('\w\d\d\d\d\d\d\d',link['href']).group())
    
    
    for url in all_link:
        
        soup = BeautifulSoup(get_requests(url),'html.parser')
        # 抓取法條區塊
        soup_content = soup.find('div',class_='law-reg-content')
        # 找到法條
        soup_by_row = soup_content.find_all('div',class_='row')
        
        
        # 抓取條號
        number = []
        # 抓取內文
        data = []
        for i in soup_by_row:
            data.append(i.find('div',class_='col-data').text)
            number.append(i.find('div',class_='col-no').text)
            
        # 抓取法規名稱
        tital = soup.find('a', id="hlLawName").text
        
        
        
        df = pd.DataFrame({
            '條號' : number,
            '內容' : data
            })
        
        df['法規名稱'] = tital

        # 加上換行符號
        df['內容']=df['內容'].apply(lambda x:x.replace('。','。\n').replace('：','：\n'))
        df = df[['法規名稱','條號','內容']]
        print(tital)
        add_to_csv(df)
  
    
def add_to_csv(df):
    final_file = os.path.join(os.path.join('datasource','base.csv'))
    old_df = pd.read_csv(final_file) if os.path.isfile(final_file) else pd.DataFrame()
    old_df = pd.concat([old_df,df])
    old_df.drop_duplicates(inplace = True)
    old_df.to_csv(os.path.join('datasource','base.csv') ,encoding = 'utf-8-sig' , index=False )

def search_keyword(search_key_word):
    
    old_df = pd.read_csv(os.path.join('datasource','base.csv'))
    # old_df['內容']=old_df['內容'].apply(lambda x:x.replace('，','\n，'))
    search = old_df['內容'].str.contains(search_key_word)
    
    file_path =os.path.join('search' , f'{search_key_word}_search.csv')
    print(file_path)
    old_df[search].to_csv(file_path,encoding = 'utf-8-sig' , index=False )
    
    


