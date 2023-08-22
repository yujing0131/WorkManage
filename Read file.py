from openpyxl import load_workbook
import pandas as pd 

input_file1="C:/Users/tcpi-q06/Downloads/上月底+本月新增.csv"
input_file2="C:/Users/tcpi-q06/Downloads/本月底+本月減少.csv"
output_folder="D:/外網統計園地/python"
names = ['01五年在監', '02入監罪名', '03入監刑名', '04入監教育', '05入監年齡', '06出獄罪名', '07假釋罪名', '08在監罪名', '09在監應執刑名', '10在監教育', '11在監年齡', '12戒治五年人數', 
'13戒治新入毒品級別', '14戒治新入教育', '15戒治新入年齡', '16戒治在所毒品級別', '17戒治在所教育', '18戒治在所年齡']


workbook1 = pd.read_csv(input_file1)
workbook2 = pd.read_csv(input_file2)
print(workbook1["收容人呼號"])
print(workbook2["收容人呼號"])
