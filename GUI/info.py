# -*- coding: utf-8 -*-
"""
Created on Fri Jan  6 15:36:02 2023

@author: LV
"""

import pandas as pd
import datetime 
import re
import numpy as np

df1 = pd.read_excel('d:/DATABASE/信息提取.xlsx')
list1=list(df1['开票信息副本'])
IDlist=[]
namelist=[]
placelist=[]
phonelist=[]
banklist=[]
bankconlist=[]
result=pd.DataFrame()
for i in list1:
    seglist=i.split("\n")
    cus_ID=re.findall(r'9[0-9A-Z]{14,17}', i)
    name=re.findall(r'名称：(.*)|(.*公司)', i)
    place=re.findall(r'址：(.*)|(海南省.*)', i)
    phone=re.findall(r'1[0-9]{10}|[0-9]{3,4}-[0-9]{8}', i)
    bank=re.findall(r'([中海].*[支分]行)', i)
    bankcon=re.findall(r'([2-9]+[0-9]{4,25}$)', i)
    try:
        IDlist.append(cus_ID[0])
    except:
        IDlist.append('')
    try:
        namelist.append(name[0])
    except:
        namelist.append('')
    try:
        placelist.append(place[0])
    except:
        placelist.append('')   
    try:
        phonelist.append(phone[0])
    except:
        phonelist.append('')
    try:
        bankconlist.append(bankcon[0])
    except:
        bankconlist.append('')
    try:
        banklist.append(bank[0])
    except:
        banklist.append('')
for q in range(len(placelist))  :
    if type(placelist[q])!=str:
        if placelist[q][0] =='':
            placelist[q]=placelist[q][1]
        else:
            placelist[q]=placelist[q][0]
for e in range(len(placelist))  :   
    if re.findall(r'1[0-9]{10}', placelist[e])!=[]:
        placelist[e]=placelist[e][0:-11]       
result['统一社会信用代码']=IDlist
result['企业名称']=namelist
result['地址']=placelist
result['电话']=phonelist
result['开户行']=banklist
result['银行账号']=bankconlist






