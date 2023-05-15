# -*- coding: utf-8 -*-
"""
Created on Tue Jun 14 15:27:24 2022

@author: LV
"""


import xlrd

import numpy as np
import pandas as pd
import datetime 
def run():
    xlsx = pd.ExcelFile('d:/analyze program/会员台账.xls')
    xlsx1 = pd.ExcelFile('d:/analyze program/洋浦纳统情况.xlsx')
    df1 = pd.read_excel('d:/analyze program/22.xlsx')
    df2 = pd.read_excel('d:/analyze program/海口纳税0112.xlsx')
    df3 = pd.read_excel('d:/analyze program/洋浦1220纳税.xlsx')
    df4 = pd.read_excel('d:/analyze program/14.xlsx')
    df5 = pd.read_excel('d:/analyze program/15.xlsx')
    
    df2['tag']=df2['公司名称']+df2['税款所属月份'].astype('str')

    
    
    df = pd.read_excel(xlsx,0)      
    df1 = pd.read_excel(xlsx1,0)    
    df2 = pd.read_excel(xlsx1,2) 
    
    result=pd.merge(df1, df2,on='企业名称',how='left')
    result2=pd.merge(df1, df3,on='统一社会信用代码',how='left')
    result3=pd.merge(result2, df4,on='公司名称',how='outer')
    result4=pd.merge(result3, df5,on='公司名称',how='outer')
    
    
    result1=pd.merge(df2, df1,on='统一社会信用代码',how='left')


   # result1=pd.merge(result, yp,on='企业名称',how='left')
    #result1=pd.merge(df, df_3,left_on=(['统一社会信用代码','税收所属期']),right_on=(['统一社会信用代码','开票月份']),how='left')
    
    result.to_excel("d:/analyze program/情况1.xlsx")
    result1.to_excel("d:/analyze program/情况2.xlsx")
    #result1.to_excel("d:/analyze program/结果.xlsx")
    #result=pd.merge(result, df6,on='公司',how='left')

