from apyori import apriori
from openpyxl import load_workbook
import pandas as pd
input_file="C:/Users/tcpi-q06/Desktop/違規規則_1120414高女監獄.xlsx"
data = pd.read_excel(input_file,sheet_name="data1(無序號)",header=0)
#print(data.iloc[1])

factor=data.iloc[:,1:]
#print(factor.iloc[1])

result= data['有違規前科無違規前科違規前科違規前科']
result1=[]
result2=[]
data1=[]
for i in range(1,len(data)):
    result[i]=result[i].replace("違規前科","")
    if result[i]=="有":
        print(factor.iloc[i])
        result1.append(list(factor.iloc[i]))
        data1.append(list(data.iloc[i]))
    if result[i]=="無":
        print(factor.iloc[i])
        result2.append(list(factor.iloc[i]))
#print(result1)

association_rules = apriori(data1, min_support=0.16, min_confidence=0.2, min_lift=3, max_length=2) 
association_results = list(association_rules)
 
for item in association_results:
   pair = item[0] 
   items = [x for x in pair]
   print("Rule: " + items[0] + " -> " + items[1])
   print("Support: " + str(item[1]))
   print("Confidence: " + str(item[2][0][2]))
   print("Lift: " + str(item[2][0][3]))
   print("=====================================")