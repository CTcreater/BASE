# -*- coding: utf-8 -*-
"""
Created on Tue Dec  6 16:06:54 2022

@author: LV
"""
import pandas as pd
import datetime 
import re


df_use=pd.read_excel('d:/analyze program/待分类.xlsx')

df_9 = pd.read_excel('d:/DATABASE/cargotype/产品字典库.xlsx')  #df9产品字典库
list_re_cargo=list(df_9['key'])
dict_type=dict(zip(df_9['key'],df_9['大类']))

cargolist=list(df_use['种类'])
cargotype=[]
for i in cargolist:

    for x in list_re_cargo:
        pattern = re.compile(r'(.*)'+x)
        result1 = pattern.search(str(i))
        if result1 != None:
            cargotype.append(dict_type.get(x))
            break
    if result1 == None:
        cargotype.append('unknow')
df_use['大类']=cargotype

