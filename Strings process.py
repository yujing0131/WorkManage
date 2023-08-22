import pandas as pd
import openpyxl
import re
data = pd.read_excel(r'.\違規資料.xlsx',sheet_name = '明細檔')
fact = data['獎懲事實']
#print(len(fact))
#print(fact[0])
#print('時' in fact[0] or '點' in fact[0])
#print(data['獎懲事實'])
for i in range(880,881):#len(fact) 
    field=''
    area=''
    layer=''
    room=''
    
    print(i)
    ##檢查是否不為空白
    if type(fact[i])==str:
        ##檢查是否有時間資訊
        clock_bool = '時' in fact[i] or '點' in fact[i]
        minute_bool= '分' in fact[i]
        index_clock1 =fact[i].find('時')
        index_clock2 =fact[i].find('點')
        index_time  =  fact[i].find(':')
        index_minute = fact[i].find('分')
        t= re.findall(r'-?\d+\.?\d*',fact[i][:index_time])
        y= re.findall(r'-?\d+\.?\d*',fact[i][index_time:])
        s = [str(s) for s in re.findall(r'-?\d+\.?\d*',fact[i][:max(index_clock1,index_clock2)])]
        m = [str(s) for s in re.findall(r'-?\d+\.?\d*',fact[i][max(index_clock1,index_clock2):])]
        print([t,y,s,m])
        if index_time > 0 and t!=[] and y!=[] :
            print([int(a)< 13 and int(a)>-1  for a in t])
            clock = t[len(t)-1]
            minute = y[0]
            if len(clock)>0 and len(minute)>0 and len(clock)<3 and len(minute)<3 :
                time = clock+':'+ minute
            else:
                time=''
        if  max(index_clock1,index_clock2) > 0 and s!=[] and m!=[]:
            if s!=[] and m!=[]:
                clock = s[len(s)-1]
                minute = m[0]
            
            if len(clock)>0 and len(minute)>0 and len(clock)<3 and len(minute)<3:
                time = clock+':'+ minute
            else:
                time=''
        if  (s==[] or m==[]) and (t==[] or y==[]):
            time=''
        
        print('time='+time)

        ##檢查是否有位置資訊
        index_place1 =fact[i].find('忠')
        index_place2 =fact[i].find('孝')
        index_place3 =fact[i].find('仁')
        index_place4 =fact[i].find('愛')
        index_place5 =fact[i].find('信')
        index_place6 =fact[i].find('義')
        index_place7 =fact[i].find('房')
        index_place8 =fact[i].find('至')
        index_place9 =fact[i].find('於')
        index_place10 =fact[i].find('在')
        
        place_bool1 = '於' in fact[i] or '在' in fact[i]
        place_bool2 = '房' in fact[i] or '場' in fact[i]
        place_bool3 = '舍' in fact[i] or '工' in fact[i] #'工場' in fact[i] or '工廠' in fact[i] 
        index_place11 =fact[i].find('舍')
        index_place12 =fact[i].find('工廠')
        index_place13 =fact[i].find('工場')
        index_place14 =fact[i].find('工')
        area_word =['舍','工場','工廠','工']
        field_word =['忠','孝','仁','愛','信','義','真','善','美','誠','禮','智','靜','靜思','明']
        print(fact[i])
        
        if place_bool1==True or place_bool3==True:
            ##若地點位置有舍/工廠/工場但沒有房
            if max(index_place11,index_place12,index_place13,index_place14) > 0  and index_place7<0: 
                if max(index_place9,index_place10) > max(index_place11,index_place12,index_place13,index_place14) and min(index_place9,index_place10)>0 :
                    result = fact[i][min(index_place9,index_place10):max(index_place12,index_place13,index_place14)]
                if max(index_place9,index_place10) > max(index_place11,index_place12,index_place13,index_place14) and min(index_place9,index_place10)<0:
                    result = fact[i][:max(index_place11,index_place12,index_place13,index_place14)]
                else :
                    result = fact[i][max(index_place9,index_place10):max(index_place11,index_place12,index_place13,index_place14)+1]
                
                print('result='+result)
                ##教區
                field_indextofind= [j in result for j in field_word]
                
                if field_indextofind != False and any(field_indextofind):
                    field = field_word[field_indextofind.index(True)]
                else:
                    field = ''
                print('field='+field)
                ##舍(工場)
                area_indextofind= [j in result for j in area_word]
                print(any(area_indextofind))
                if any(area_indextofind):
                    area = result[result.find(field):]#max(area_indextofind.find(True))
                if any(area_indextofind)==False: #and type(data['主表_工場'][i])==str:
                    area = ''
                    #area = data['主表_工場'][i]
                print('area='+area)
                layer = ''
                room = ''
            ##若地點位置包含舍/工廠/工場且包含房
            if index_place7 > 0 :
                if max(index_place8,index_place9,index_place10) > index_place7 and min(index_place8,index_place9,index_place10)>0 :
                    result = fact[i][min(index_place9,index_place10)+1:index_place7]
                if max(index_place8,index_place9,index_place10) < index_place7 and min(index_place8,index_place9,index_place10)<0 :
                    result = fact[i][max(index_place8,index_place9,index_place10):index_place7]
                else:
                    result = fact[i][:index_place7]
                print('result='+result)
                ##教區
                field_indextofind= [j in result for j in field_word]
                if field_indextofind != False and any(field_indextofind):
                    field = field_word[field_indextofind.index(True)]
                else:
                    field = ''
                wordtoreplace = ['一','二','三','四','五','六','七','八','九','十','十一','十二']
                numberreplace = ['1','2','3','4','5','6','7','8','9','10','11','12']
                #print([wordtoreplace[i] in result for i in range(0,len(wordtoreplace))])
                replace = [wordtoreplace[i] in result for i in range(0,len(wordtoreplace))]
                #print(result)
                print('field='+field)
            
                ##需要將中文數字轉換為阿拉伯數字
                if replace != False and any(replace):
                    replace_index = replace.index(True)
                    result = result.replace(wordtoreplace[replace_index],numberreplace[replace_index])
                #print('result='+ result)
                ##舍(工場)擷取內容
                area_indextofind= [j in result for j in area_word]
                print('area_indextofind='+str(area_indextofind))
                if area_indextofind!=False and any(area_indextofind):
                    areacontent = result[result.find(field):]#max(area_indextofind.find(True))
                else:
                    areacontent = result
                print('areacontent='+areacontent)
                ##樓層資訊
                if re.findall(r'-?\d+\.?\d*',areacontent)!= []:
                    l= re.findall(r'-?\d+\.?\d*',areacontent) #[areacontent.find(field)-1:areacontent.find('房')]
                    print(l)
                    layer=l[0]

                    ##避免樓層號碼為同學呼號
                    if len(layer)>3 :
                        layer=''
                else:
                    layer=''
                    #result = result.replace(layer[0],'/')
                print('layer='+layer)
            
                ##舍(工場)
                area_indextofind= [j in areacontent for j in area_word]
                #print(area_indextofind!=False and any(area_indextofind))
                if any(area_indextofind)==True:
                    print(areacontent[areacontent.find(layer)+1:areacontent.find(area_word[area_indextofind.index(True)])])
                    area = areacontent[areacontent.find(layer)+1:areacontent.find(area_word[area_indextofind.index(True)])]#max(area_indextofind.find(True))
                else:
                    area = ''
                
                print('area='+area)
                ##房號資訊           
                #print(re.findall(r'-?\d+\.?\d*',areacontent)) 
                if  re.findall(r'-?\d+\.?\d*',areacontent)!=[]:
                    r = re.findall(r'-?\d+\.?\d*',areacontent)
                    if r!=[]:
                        room = r[len(r)-1]
                        print(room)
                    ##避免房間號碼為同學呼號
                    if len(room)>3 :
                        room=''
                else:
                    room =''
                print('room='+room)
    ##若獎懲事實欄為空白沒有內容，則參考主表_工場欄位
    else:
        area=''#data['主表_工場'][i]
        room=''
        field=''
        room=''
        #field_word =['忠','孝','仁','愛','信','義','真','善','美','誠','禮','智','靜','靜思','明']
        #field_indextofind = [i in area for i in field_word]
        #if any(field_indextofind):
            #field=field_word[field_indextofind.index(True)]
            #layer=re.findall(r'-?\d+\.?\d*',area)[0]
  

    data['違規時間'][i]=time
    data['違規地點_教區'][i]=field
    data['違規地點_樓層'][i]=layer
    data['違規地點_舍(工場)'][i]=area
    data['違規地點_房號'][i]=room
  

#print(data['違規時間'])
#print(data['違規地點_教區'])
#print(data['違規地點_樓層'])
#print(data['違規地點_舍(工場)'])
#print(data['違規地點_房號'])

#data.to_excel('違規處理_04.xlsx', sheet_name='明細檔')




