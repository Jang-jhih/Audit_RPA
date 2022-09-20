from autoWP.autoWP import *

#%%
path = r'\\192.168.248.10\IA_Evidence\FD'
# path = r'\\192.168.248.10\ia\全支付各單位辦法'

#%%
# 取出紀錄表，如果沒有就是還沒放檔案
All_AuditItem = PrintFilePath(path)

print(f'目前檔案共{len(All_AuditItem)}筆')
if len(All_AuditItem)== 0:
    print('該單位還沒放任何資料。')

else:
    MadeWP(All_AuditItem,path)

    
print(f'執行完成')


