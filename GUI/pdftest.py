# -*- coding: utf-8 -*-
"""
Created on Thu Nov  3 16:08:21 2022

@author: LV
"""

import pdfplumber
#import pymupdf
import re
import pandas as pd
import newCUSRP 
import os

# 遍历文件夹下的所谓pdf，看完税证明编号有无重复，无重复开始下面进程，并修改pdf文件名


def run(place=''):
    if place == '海口':
        drname='D:/DATABASE/training_HK/' 
        
    if place == '洋浦':
        drname='D:/DATABASE/training_YP/' 
    file_list=os.listdir(drname)

    result=pd.DataFrame()
    data=newCUSRP.initialize()
    readed_id=[]

    for i in file_list:
       
        dir_name=drname+i
    
        with pdfplumber.open(dir_name) as pdf:
            first_page =pdf.pages[0]
            

            tax_data=first_page.extract_text().split("\n")
            tax_ID=re.findall('证明 (.*)',tax_data[2])[0]
            if tax_ID not in readed_id:
                readed_id.append(tax_ID)
                
                
                
                
                cus_ID=re.findall('纳税人识别号 (.*)',tax_data[4])[0]
                cus_name=data.namedict.get(cus_ID)
                
 
                
                tax_list=tax_data[6:]
                for sn in range(len(tax_list)):
                    s=tax_list[sn]
                    s0=tax_list[sn-1]
                    tax_series=s.split(" ")
                    taxname=tax_series[0]
                    if taxname in ['印花税','增值税','城市维护建设税','教育费附加','地方教育附加','企业所得税','个人所得税']:
                        tax_start=tax_series[1]
                        tax_end=tax_series[3]
                        tax_putdate=tax_series[4]
                        tax_amount=tax_series[5]
                        try :
                            tax_region=tax_series[6]
                        except:
                            tax_region=''
                        tax_serise = pd.Series()
                        tax_serise['完税证明编号']=tax_ID
                        tax_serise['统一社会信用代码']=cus_ID
                        tax_serise['税种']=taxname
                        tax_serise['起始日']=tax_start
                        tax_serise['截止日']=tax_end
                        tax_serise['税款入库日期']=tax_putdate
                        tax_serise['入库地']=tax_region
                        tax_serise['金额']=tax_amount
                        result=result.append(tax_serise,ignore_index=True)  
                 
            else:
                print(tax_ID+'已经有了!!!')
        new_file_name=cus_ID+'-'+tax_ID+'.pdf'
        
        if i != new_file_name:
            try:
                os.rename(dir_name, drname+new_file_name)
                print(i+'重命名完成')
            except:
                os.rename(dir_name, drname+new_file_name+'重复')
                print(i+'重命名完成')
        else:
            pass
                        
    return result
    

    
    
