# -*- coding: utf-8 -*-
"""
Created on Tue Oct 25 10:06:01 2022

@author: LV
"""

import xlwings as xw
import os
import datetime
import pandas as pd
import newCUSRP 
from decimal import *


# 加载 excel 文件

def run(start=202201,end=202211,drop_applyed=False,makesure=False,confirm=False,usedata='raw'):
    

    
    file_dir='d:/analyze program/洋浦返税材料0105'
    now=datetime.datetime.now()
    today=now.strftime('%Y年%m月%d日')
    Filename='/'+str(start)+'--'+str(end)
    os.makedirs(file_dir+Filename)
    data=newCUSRP.initialize()
    app=xw.App(visible=False,add_book=False)
    df_10=data.df_10
    df_10.fillna(1,inplace=True)
    name_df=df_10[df_10['官方类型']!=1]
    name_df=name_df[name_df['编码']!=1]
    code_dict=dict(zip(name_df['统一社会信用代码'],name_df['编码'].astype('str')))
    name_df['注册或迁址到洋浦时间']=name_df['注册或迁址到洋浦时间'].apply(lambda x :x.strftime('%Y年%m月%d日'))    
    regdate_dict=dict(zip(name_df['统一社会信用代码'],name_df['注册或迁址到洋浦时间']))
    name_key=list(name_df['统一社会信用代码'])
    if usedata=='new':
        data.tax_remake()
    raw_table=data.yp_apply(need='new',start=start,end=end)
    raw_table.to_excel(file_dir+Filename+'/汇总明细单.xlsx',index=False)
    raw_table=raw_table[raw_table['是否保留']=='yes']
    df1=data.df_1
    stop_dict=dict(zip(df1['统一社会信用代码'],df1['停止合作时间']))
    
    
    # 得到sheet对象
    
    
    lost=pd.DataFrame()
    apply_df=pd.DataFrame()
    
    
    for i in name_key:
        temp=raw_table[raw_table['统一社会信用代码_x']==i]
        if drop_applyed == True:
            applyed_df=data.df_yp_applyed
            applyed_df['标记']=applyed_df['统一社会信用代码_x']+applyed_df['税收所属期'].astype('str')
            marklsit=list(applyed_df['标记'])
            filter_condition={'标记':marklsit}
            temp['标记']=temp['统一社会信用代码_x']+temp['税收所属期'].astype('str')
            temp=temp[~temp.isin(filter_condition)['标记']]
            temp.drop(columns=['标记'],inplace=True)

        
        
        
        enddate=stop_dict.get(i)
        try:
            end_line=int(enddate[:4]+enddate[5:7])
            temp=temp[temp['税收所属期']<end_line]  
            print(i+'在'+str(end_line)+'已停止合作')
            
        except:pass
        sale_dict=dict(zip(temp['税收所属期'],temp['价税合计']))
        apply_dict=dict(zip(temp['税收所属期'],temp['营业收入1.5%']+temp['营业收入1.25%']))
        
        
        wb = app.books.open('d:/DATABASE/洋浦模板.xlsx')
        sheet = wb.sheets(1)
        #1月价税合计
        sheet.range("A2").value = '企业名称（盖章）：'+data.namedict.get(i)
        sheet.range("A3").value = '在洋浦注册或迁入洋浦日期：'+regdate_dict.get(i)
        sheet.range("K3").value = '编号：'+code_dict.get(i)
        
        
        write_dict1={202201:"C6",202202:"F6",202203:"I6",202204:"L6",202205:"C7",202206:"F7",202207:"I7",202208:"L7",202209:"C8",202210:"F8",202211:"I8",202212:"L8"}
        write_dict2={202201:"C10",202202:"F10",202203:"I10",202204:"L10",202205:"C11",202206:"F11",202207:"I11",202208:"L11",202209:"C12",202210:"F12",202211:"I12",202212:"L12"}
        
        
       
        
        monthrange=list(range(start,end+1))
        resum=0.00
        resum2=0.00
        for y in monthrange:
            coordinate=write_dict1[y]
            
            try:
                seg_sum=sale_dict.get(y)+0.000001
                sheet.range(coordinate).value = '%.2f'%seg_sum+'元'
                n=round(sale_dict[y]+0.000000001,2)
            except:
                sheet.range(coordinate).value = '0.00元'  
                n=0 
            resum+=n    
            
            coordinate2=write_dict2[y]
            
            
            try:
                seg_sum2=apply_dict.get(y)+0.000001
                sheet.range(coordinate2).value = '%.2f'%seg_sum2+'元'
                n2=round(apply_dict[y]+0.000000001,2)
            except:
                sheet.range(coordinate2).value = '0.00元' 
                n2=0
                
            resum2+=n2
        sheet.range("B9").value = '%.2f'%resum+'元'
        sheet.range("B13").value = '%.2f'%resum2+'元'      
       
        
      #求和要求每个月先四舍五入，再求和相加  

      
        
        
        
        
        e1=1
        for e in range(10):
            try:
                q=sale_dict[202201+e]
            except:
                q=0  
            e1*=q

        
        

        
        if makesure == True: 
            if e1 !=0:
                apply_df=apply_df.append(temp,  ignore_index=True)
                wb.save('d:/analyze program/洋浦返税材料0105'+Filename+'/'+data.namedict.get(i)+'.xlsx')
                
            else:
                print (i+'有缺失')
                lost=lost.append(temp,  ignore_index=True)
        else:
            
            apply_df=apply_df.append(temp,  ignore_index=True)
            try:
                wb.save('d:/analyze program/洋浦返税材料0105'+Filename+'/'+data.namedict.get(i)+'.xlsx')
            except:
                print (i+'发生异常')
            print (i+'有缺失')
            lost=lost.append(temp,  ignore_index=True) 
            
        
        wb.close()
        
    lost=data.add_tax_person(raw_pd=lost,left_on='统一社会信用代码_x',on='统一社会信用代码')
    lost.to_excel('d:/analyze program/洋浦返税材料/缺失企业.xlsx',index=False)
    if confirm == True:
        fname=today+'apply.xlsx'
        fname2=today+'apply_all.xlsx'
        


        apply_df['申请日期']=data.today.strftime("%Y-%m-%d %H:%M:%S")
        apply_df.to_excel('d:/DATABASE/applyed/yp/'+fname,index=False)

        newapplyed=data.df_yp_applyed.append(apply_df,  ignore_index=True)
        newapplyed.to_excel('d:/DATABASE/applyed/yp/'+fname2,index=False)
        newapplyed.to_excel('d:/DATABASE/all/ypapplyed.xlsx',index=False)
        
        
    app.quit()
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
