# -*- coding: utf-8 -*-
"""
Created on Tue Sep  6 15:24:39 2022

@author: LV
"""

import newCUSRP as cup
import openpyxl
import datetime
import pandas as pd
from datetime import timedelta
data=cup.initialize()
# 加载 excel 文件
wb = openpyxl.load_workbook('d:\DATABASE\每日汇报.xlsx')
now=datetime.datetime.now()
today=now.strftime('%Y%m%d')
# 得到sheet对象
sheet = wb['数据']
num1=6
for i in ['2022','2021']:
    
    if i =='累计':
        start=202001
        end=data.tomonth
    else:
        start=int(i+'01')
        end=int(i+'12')
    for y in ['海口','洋浦','']:
        d1=data.report(start,end,y,'sale')
        p1='E'+str(num1)
        sheet[p1] = d1.sum()['价税合计']
        d2=data.report(start,end,y,'tax')
        p2='F'+str(num1)
        sheet[p2] = d2.sum()['合计']
        d3=data.report(start,end,y,'taxback_real')
        p3='G'+str(num1)
        sheet[p3] = d3.sum()['实返合计']
        
        
        #b1=data.boss(start,end,y,'sale')
        #b2=data.boss(start,end,y,'tax')
       # b3=data.boss(start,end,y,'taxback_real')
        
        num1+=2
    num1+=1
    
data.tax_refresh()
noback_hk=data.TaxBack_calculate(start=202201,place='海口')
q1=noback_hk.sum()['申请合计']
noback_hk_1=noback_hk[noback_hk['申请合计']!=0]
q2=noback_hk_1.shape[0]

noback_yp=data.TaxBack_calculate(start=202201,place='洋浦') 
q3=noback_yp.sum()['申请合计'] 
noback_yp_1=noback_yp[noback_yp['申请合计']!=0]
q4=noback_yp_1.shape[0]
        
sheet['H6'] =  q1 
sheet['I6'] =  q2      
sheet['H8'] =  q3
sheet['I8'] =  q4


result=pd.merge(data.df_11, data.df_1,left_on='企业名称',right_on='企业名称',how='left')
use_temp=result[result['第一年']>='2022-01-01']
use_hk=use_temp[use_temp['注册地_x']=='海口']
use_yp=use_temp[use_temp['注册地_x']=='洋浦']
sheet['C5'] =  use_hk.shape[0]
sheet['C7'] =  use_yp.shape[0]

now=data.today-timedelta(days=7)
now_str=now.strftime('%Y-%m-%d')
use_temp_week=result[result['第一年']>=now_str]
use_hk_week=use_temp_week[use_temp_week['注册地_x']=='海口']
use_yp_week=use_temp_week[use_temp_week['注册地_x']=='洋浦']
sheet['C2'] =  use_hk_week.shape[0]
sheet['C3'] =  use_yp_week.shape[0]

now2=datetime.date(now.year,now.month, 1)
now_str=now2.strftime('%Y-%m-%d')
use_temp_month=result[result['第一年']>=now_str]
use_hk_month=use_temp_week[use_temp_month['注册地_x']=='海口']
use_yp_month=use_temp_week[use_temp_month['注册地_x']=='洋浦']
sheet['C33'] =  use_hk_month.shape[0]
sheet['C35'] =  use_yp_month.shape[0]

## 指定不同的文件名，可以另存为别的文件
filename=today+'.xlsx'
wb.save('d:\DATABASE\每日汇报'+filename)

