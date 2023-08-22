import pandas as pd
import os.path
#from asposecells import workbook
from pyexcel_ods3 import save_data
from pyexcel_xlsx import get_data
from collections import OrderedDict
from openpyxl import load_workbook
  

file = "統計園地上網.xlsm"
names = ['01五年在監', '02入監罪名', '03入監刑名', '04入監教育', '05入監年齡', '06出獄罪名', '07假釋罪名', '08在監罪名', '09在監應執刑名', '10在監教育', '11在監年齡', '12戒治五年人數', 
'13戒治新入毒品級別', '14戒治新入教育', '15戒治新入年齡', '16戒治在所毒品級別', '17戒治在所教育', '18戒治在所年齡']
sheet = pd.read_excel(file, header=None, sheet_name=names) #sheet_name=None表示讀取所有表 
#print(sheet)
for k in range(0,len(names)):
    data = pd.DataFrame(sheet[names[k]])
    final= pd.DataFrame()#columns=len(data.columns),index = range(0,len(data.columns))

    for i in range(0,len(data.columns)):#
        finaldata=[]
        for j in range(0,len(data[i])):
            series= data[i][[j]].values[0]
            
            if type(series)==float:
                series =''
                element =series
            if type(series)!=str:
                element = str(series)
            else:
                element = series
            #print(type(element))
            finaldata.append(element)
        final[i] = finaldata
        #print(final)
    result = OrderedDict()
    result.update({ names[k]: final.to_numpy().tolist()})
    print(result)
    
    #directory = 'D:\外網統計園地\11201'
    filename = names[k]+".ods"
    save_data(filename,result)
    #file_path = os.path.join(directory, filename)
    #if not os.path.isdir(directory):
        #os.mkdir(directory)
    #with open(os.path.join(directory,filename), "w") as file:
        #file.write(result[str(names[k])])
    


