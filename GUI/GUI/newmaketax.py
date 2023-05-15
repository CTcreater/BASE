# -*- coding: utf-8 -*-
"""
Created on Fri Jan  6 17:36:03 2023

@author: LV
"""

import os
import xlwings as xw
import pandas as pd
#from interval import Interval
import warnings
import newCUSRP 
import wordkiller as wk
import commitment
import numpy as np


#'d:/analyze program/see.xlsx'
warnings.filterwarnings('ignore')

def make_material(from_date,end_date,place='海口',confirm=False,drop_applyed=False,newtax=True):
    file_dir='d:/analyze program/'
    Filename=place+str(from_date)+'至'+str(end_date)+'新返税材料包'
    data=newCUSRP.initialize()
    if newtax == True:
        data.tax_remake()
        print('纳税数据已切换至新版')
    
    def round1(x,y=2):
        if x>=0:
            return round(x+0.0000001,y)
        else :
            return round(x-0.0000001,y)
    
    
    app=xw.App(visible=False,add_book=False)
    df1=data.df_1
    stop_dict=dict(zip(df1['统一社会信用代码'],df1['停止合作时间']))
    
   
    if place == '海口':
    
 
        prove_df=data.df_81  
        prove_df=prove_df.drop(columns=['Unnamed: 26','Unnamed: 36','Unnamed: 37','Unnamed: 38'
                               ,'Unnamed: 39','Unnamed: 40','Unnamed: 41','Unnamed: 42'
                               ,'Unnamed: 43','Unnamed: 44','Unnamed: 45'])
        star=str(from_date)[:4]+"/"+str(from_date)[4:6]+"/01"
        end=str(end_date)[:4]+"/"+str(end_date)[4:6]+"/31"
        mask=(prove_df['Unnamed: 13'] >= star) & (prove_df['Unnamed: 13'] <= end)      
        userange=prove_df.loc[mask]   #总的完税计算表合集
        
        

    
    
        
        
        
        #刷新userange，丢弃超过停止合作期间入库的税款
        userange=data.add_end_date(userange, 1,on='Unnamed: 5')
        userange['税款是否过服务期']=np.where(userange['Unnamed: 13']>userange['停止合作时间'],1,0)
        userange=userange[userange['税款是否过服务期']!= 1 ]
        
        #刷新userange，丢弃超过税款截止日在签订日期之前的税款
      

        
        if confirm == True:
            userange.to_excel('d:/DATABASE/applyed/海口已申请.xlsx',index=False)
        customer_set=set(list(userange['Unnamed: 5']))
        
        os.makedirs(file_dir+Filename)
        
        

        tax_sum=data.tax_raw_cube(DF=userange)   #drop_close 参数：丢弃超过停止合作期限的税款 
        
        
        
        seer=tax_sum.groupby('统一社会信用代码').sum()
        seer['纳税合计']=seer['增值税']+seer['城建税']+seer['教育费附加']+seer['地方教育附加']+seer['印花税']+seer['企业所得税']+seer['个人所得税']+seer['其他收入-工会经费']
        seer.reset_index(inplace=True)
        seer.drop(columns=['税收所属期','Unnamed: 0','Unnamed: 1'],inplace=True)
        seer=data.add_name(seer, 1)
        seer=data.add_last_renew(seer, 1)
        seer=data.add_serve_start(seer, 1)
        wk.make(df=seer, start=from_date, end=end_date, path=file_dir+Filename+'/'+'运营总部关于申请财政奖励的说明.docx')
        
        
        
        
        
        seer=data.add_end_date(seer, 1,on='统一社会信用代码')
        
        seer.to_excel(file_dir+Filename+'/'+'可申请奖励企业详情.xlsx',index=False)
        
        #完税证明丢弃掉无用字段
        userange_need=userange.loc[:,['合计','省级财力贡献','市级财力贡献']]
        userange_drop=userange.drop(columns=['Unnamed: 0','Unnamed: 1','Unnamed: 7','Unnamed: 14','Unnamed: 15','Unnamed: 25','增值税.1','城建税.1','印花税.1','企业所得税.1','个人所得税.1','合计','省级财力贡献','市级财力贡献'
                                             ,'Unnamed: 32','税款是否过服务期'])
        userange_drop.fillna(0, inplace=True)
  #      list_time=['' for i in range(len(userange_drop))]
 #       userange_drop.insert(8,'申请时间',list_time)
        userange_drop['累计纳税']=userange_drop['增值税']+userange_drop['城建税']+userange_drop['教育费附加']+userange_drop['地方教育附加']+userange_drop['印花税']+userange_drop['企业所得税']+userange_drop['个人所得税']+userange_drop['其他收入-工会经费']
        userange_drop['平衡-企业所得税']=userange_drop['企业所得税']*0.34
        userange_drop['平衡-增值税']=userange_drop['增值税']*0.425
        userange_drop['平衡-城建税']=userange_drop['城建税']*0.85
        userange_drop['平衡-印花税']=userange_drop['印花税']*0.85
       # userange_drop['平衡-个人所得税']=userange_drop['个人所得税']*0.4
        userange_drop['合计']=userange_need['合计']
        userange_drop['省级财力贡献']=userange_need['省级财力贡献']
        userange_drop['市级财力贡献']=userange_need['市级财力贡献']
        userange_drop.insert(5,'序号',userange_drop.index)
        userange_drop.insert(10,'是否被列入异常经营企业名录','')
        userange_drop.insert(11,'列异起止时间','')
        userange_drop.insert(12,'是否递交承诺书','')                              
        userange_drop_show=userange_drop.iloc[:,:25]     
        df_3=data.df_3     
        df_3['销方企业名称']=df_3['统一社会信用代码'].apply(lambda x :data.namedict.get(x))              
                                             
        for i in  customer_set:
            
            
            enddate=stop_dict.get(i)
            
            
            
            

            cus_name=data.namedict.get(i)
            os.makedirs(file_dir+Filename+'/'+cus_name)
            temp=userange[userange['Unnamed: 5']==i]
            try:
                end_line=int(enddate[:4]+enddate[5:7])
                temp=temp[temp['Unnamed: 13']<end_line]
                print(data.namedict(i)+'在'+str(end_line)+'已停止合作')
                
            except:pass
            
            
            



            
            wb_2=app.books.open('d:/DATABASE/海口0106模板.xlsx')
            sh2_1 = wb_2.sheets(1)
            sh2_2 = wb_2.sheets(2)
            temp2=userange_drop_show[userange_drop_show['Unnamed: 5']==i]
        
            lentemp=len(temp2)+7
            seg1='A'+str(lentemp)+':AG1120'
            try:
                end_line=int(enddate[:4]+enddate[5:7])
                temp2=temp2[temp2['Unnamed: 13']<end_line]
                
                
            except:pass
            
            
            
            
            sh2_2.range("A7:Y1120").value = temp2.values
            sh2_2[seg1].delete()
            
            lenth=(int(str(end_date)[:4])-int(str(from_date)[:4]))*12-int(str(from_date)[4:6])+int(str(end_date)[4:6])+1
            
            list_ym=[]
            for n in range(lenth):
                if n+int(str(from_date)[5:6])>12:
                    seg=from_date+n-12+100
                    list_ym.append(seg)
                else:
                    list_ym.append(from_date+n)
                    
            temp1=pd.DataFrame()
            temp1['企业名称']=[data.namedict.get(i) for q in range(lenth)]
            temp1['年月']=list_ym
            temp11=pd.merge(temp1, df_3,left_on=['企业名称','年月'],right_on=['销方企业名称','开票月份'],how='left')
            temp1_use=temp11.loc[:,['企业名称','年月','金额']]
            temp1_use['金额2']=temp1_use['金额']
            temp1_use1=temp1_use.loc[:,['金额','金额2']]
            sh2_1.range("C5:D16").value =temp1_use1.values
            
            userange_drop_temp=userange_drop[userange_drop['Unnamed: 5']==i]
            userange_drop_temp['key']=userange_drop_temp['Unnamed: 13'].apply(lambda x :x[:4]+x[5:7])

            s1=userange_drop_temp.groupby('key').sum()
            dict_all=s1.to_dict('index')
            
            
            tax_sum=[]
            tax_back=[]
            
            for x in list_ym:
                try:
                    tax_sum.append(dict_all[str(x)]['累计纳税'])
                except:
                    tax_sum.append('')
                    
                try:
                    tax_back.append(dict_all[str(x)]['合计'])
                except:
                    tax_back.append('')   
            temp12=pd.DataFrame()
            temp12['A']=tax_sum
       
            temp12['C']=tax_sum
     
            sh2_1.range("F5:G16").value = temp12.values

            wb_2.save(file_dir+Filename+'/'+cus_name+'/'+cus_name+'地方财政奖励计算表.xlsx')
            wb_2.close()   
            now=data.today.strftime('%Y年%m月%m日')
            path1=file_dir+Filename+'/'+cus_name+'/'+cus_name+'承诺书.docx'
            commitment.make(name=cus_name, Date='', path=path1)
        
        
        app.quit()
        
