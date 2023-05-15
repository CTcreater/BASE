# -*- coding: utf-8 -*-
"""
Created on Fri Jan 13 16:31:59 2023

@author: LV
"""

import pandas as pd
import datetime 
import re
import numpy as np

df1 = pd.read_excel('d:/DATABASE/未收到的协议.xlsx')
df2 = pd.read_excel('d:/DATABASE/all/会员台账.xls')
list1=list(df1['企业名称'])
list2=list(df2['企业名称'])
list3=list2[1:]
IDlist=[]
namelist=[]
for i in range(len(list1)):

    for e in range(len(list3)):
        full_name=re.findall(r'(.*)'+list1[i]+r'(.*)', list3[e])
        if full_name!=[]:
            print(list3[e])
            list1[i]=list3[e]
        
df1['全名']=list1




