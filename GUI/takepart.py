# -*- coding: utf-8 -*-
"""
Created on Mon Oct 17 10:56:09 2022

@author: LV
"""
import xlwings as xw

import datetime
import pandas as pd


# 加载 excel 文件

def run():
    app=xw.App(visible=False,add_book=False)
    yh=pd.read_excel('d:/拆分工具/招商机构银行账号.xlsx')
    jsd=pd.read_excel('d:/拆分工具/招商机构结算单.xlsx',header=1)
    jsd.drop([len(jsd)-1],inplace=True)
    jsd.drop([len(jsd)-1],inplace=True)
    
    
    now=datetime.datetime.now()
    today=now.strftime('%Y年%m月%d日')
    # 得到sheet对象
    
    
    
    
    jg_list=list(set(jsd['招商机构']))
    zhanghudict=dict(zip(yh['招商机构收款人全称'],yh['收款账号']))
    kaihudict=dict(zip(yh['招商机构收款人全称'],yh['开户行']))
    for i in jg_list:
        
        
        
        tempuse=jsd[jsd['招商机构']==i]
        shou=list(set(tempuse['收款单位']))
        
        for x in shou:
            temp=tempuse[tempuse['收款单位']==x]
            wb = app.books.open('d:\拆分工具\模板.xlsx')
            sheet = wb.sheets(1)
            sheet.range("A8:H59").value = temp.values[:,2:]
            sheet.range("C2").value =x
            sheet.range("C3").value =i
            sheet.range("C4").value =i
            sheet.range("C5").value =zhanghudict.get(i)
            sheet.range("C6").value =kaihudict.get(i)
            print(i+str(zhanghudict.get(i))+'_____'+str(kaihudict.get(i)))
            sheet.range("D66").value =i
            
            
            sheet.range("B73").value =x+'开票信息'
            sheet.range("B74").value ='单位名称：'+x
            
            if x=='海南国际能源交易中心有限公司':
                sheet.range("B75").value ='税号：91460000MA5TAU21X3'
                sheet.range("B76").value ='单位地址：海南省洋浦经济开发区控股大道1号洋浦迎宾馆裙楼'
                sheet.range("B77").value ='电话号码：0898-68638263'
                sheet.range("B78").value ='开户银行：中国建设银行股份有限公司海口海府支行'
                sheet.range("B79").value ='银行账号：4605 0100 2236 0998 8888'
            if x=='海南国际能源交易中心运营总部有限公司':
                sheet.range("B75").value ='税号：91460100MA5TB9NX3U'
                sheet.range("B76").value ='单位地址：海南省海口市江东新区江东大道200号海南能源交易大厦1701A户'
                sheet.range("B77").value ='电话号码：0898-68638263'
                sheet.range("B78").value ='开户银行：中国银行股份有限公司海口金贸支行'
                sheet.range("B79").value ='银行账号：2650 3247 2374'
            lentemp=len(temp)+8
            seg='A'+str(len(temp)+8)+':H59'
            if x=='海南国际能源交易中心有限公司':
                x1='--交易中心'
            if x=='海南国际能源交易中心运营总部有限公司':
                x1='--运营总部'
            sheet[seg].delete()
            filename=i+x1+'.xlsx'
            
            wb.save('d:/拆分工具/12月out/'+filename)
            
            wb.close()
    
    app.quit()
    
