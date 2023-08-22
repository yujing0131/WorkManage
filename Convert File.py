#import  jpype
#import asposecells      
#jpype.startJVM() 
#from asposecells.api import Workbook
import pandas as pd
import tabula
import PyPDF2

# opened file as reading (r) in binary (b) mode
input_pdf = "D://外網統計園地//重罪不得假釋名冊//" + "1120726重罪累犯不得假釋名冊.pdf"
output_pdf = "D://外網統計園地//重罪不得假釋名冊//" + "1120726重罪累犯不得假釋名冊.xlsx"

file = open(input_pdf,'rb')
  
# store data in pdfReader
tables = PyPDF2.PdfReader(file)
  
# count number of pages
totalPages = len(tables.pages)
print(totalPages)
#names = ['01五年在監', '02入監罪名', '03入監刑名', '04入監教育', '05入監年齡', '06出獄罪名', '07假釋罪名', '08在監罪名', '09在監應執刑名', '10在監教育', '11在監年齡', '12戒治五年人數', '13戒治新入毒品級別', '14戒治新入教育', '15戒治新入年齡', '16戒治在所毒品級別', '17戒治在所教育', '18戒治在所年齡']
# Read a PDF File table
df = pd.DataFrame()
for i in range(1,totalPages+1):
    #pandas_options={'header': None} is used not to take first row as header in the dataframe.
    df_new = pd.DataFrame(tabula.io.read_pdf(input_pdf, pages=i,pandas_options={'header': None})[0])
    for j in range(0,len(df_new)):
        df_new.iloc[j,0]= str(df_new.iloc[j,0])
        df_new.iloc[j,1]= str(df_new.iloc[j,1])
    df = pd.concat([df, df_new], axis=0)

#print(df)
with pd.ExcelWriter(output_pdf) as writer: 
    df.to_excel(writer, index=False,header=False,engine='openpyxl')


# Convert into Excel File
#df.to_excel("D://外網統計園地//重罪不得假釋名冊","1120726重罪累犯不得假釋名冊.xlsx")

#將xlsm轉換成xlsx和ods檔案，這是以前寫的
#workbook = Workbook("統計園地上網.xlsm")
#workbook.save("統計園地上網.ods")
#workbook.save("統計園地上網.xlsx")
#workbook.save("統計園地上網.pdf")
#jpype.shutdownJVM()


