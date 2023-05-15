# -*- coding: utf-8 -*-
"""
Created on Wed Jul 13 13:47:00 2022
@author: LV
"""
import requests
import re
import urllib
import datetime
import numpy as np
import pandas as pd
import calendar
import warnings
import time
warnings.filterwarnings('ignore')
def autoDownload():
    now=datetime.datetime.now()
    today=now.strftime('%Y%m%d')
    tomonth=now.strftime('%Y%m')
    #登陆获取cookies
    url_login='https://www.hntradesp.com/ocrm/login'
    Payload_login={
        'username':'lvchangting',
        'password': 'Lht893104.',
        'rememberMe': 'false'
        }
    response_login=requests.post(url_login,data=Payload_login,verify=False)
    cookies_login =re.findall("'Set-Cookie': '(.*?);", str(response_login.headers))
    
    
    #完税证明新的导出方法
    url_new_prove='https://www.hntradesp.com/ocrm//tax/taxPaymentAccount/asyncExport.ajax'
    Payload_prove_hk={
        'taxRegCity':'HK'   
        }   
    head_prove_hk={'Cookie': cookies_login[0],'Accept-Encoding':''}
    response_prove_hk=requests.post(url_new_prove,data=Payload_prove_hk,headers=head_prove_hk,verify=False)
    print(response_prove_hk.text)
    for e in range(10):
        
    
        time.sleep(1)
        print(e)
    url_list='https://www.hntradesp.com/ocrm/info/fileInfo/list'
    Payload_list={
        
        'pageSize': "10",
        'pageNum' : '1'
        
        }
    response_list=requests.post(url_list,headers=head_prove_hk,data=Payload_list,verify=False)
    
    url_HK_download=re.findall('"fileUrl":"(.*?)"',response_list.text)[0]
    file_info_hk = requests.get(url_HK_download,headers=head_prove_hk,verify=False)
    
    with open(r'd:/DATABASE/Taxproof/'+now.strftime('%Y-%m-%d')+'海口完税证明.xlsx', 'wb') as file:
        file.write(file_info_hk.content)
    with open(r'd:/DATABASE/all/'+'海口完税证明.xlsx', 'wb') as file:
        file.write(file_info_hk.content)
    print('2021年10月'+'至'+now.strftime('%Y-%m')+"海口完税证明导出完成...")   
    
    # 返税数据新的导出方法
    url_tax_back='https://www.hntradesp.com/ocrm/tax/rebate/export.ajax'
    Payload_tax_back={
        
        }   
    head_tax_back={'Cookie': cookies_login[0],'Accept-Encoding':''}
    response_tax_back=requests.post(url_tax_back,data=Payload_tax_back,headers=head_tax_back,verify=False)
    print(response_tax_back.text)
    print('将等待10s')
    time.sleep(15)
            #request finished,import next
           
    url_list='https://www.hntradesp.com/ocrm/info/fileInfo/list'
    Payload_list1={
        
        'pageSize': "25",
        'pageNum' : '1'
        
        }
    response_list1=requests.post(url_list,headers=head_tax_back,data=Payload_list1,verify=False)
    
    url_taxback_download=re.findall('"fileUrl":"(.*?)"',response_list1.text)[0]
    file_info_taxback = requests.get(url_taxback_download,headers=head_tax_back,verify=False)
    
    with open(r'd:/DATABASE/Taxback/'+now.strftime('%Y-%m-%d')+'返税.xls', 'wb') as file:
        file.write(file_info_taxback.content)
    with open(r'd:/DATABASE/all/'+'返税.xls', 'wb') as file:
        file.write(file_info_taxback.content)
    print('返税数据导出完成...')   
    


  #  url_new_prove_download='https://www.hntradesp.com/ocrm//tax/taxPaymentAccount/asyncExport.ajax'
 #   response_prove_hk_download=requests.post(url_new_prove_download,headers=head_prove_hk)
    #开票数据统计按企业统计（按时间筛选）
    url_salesStatistics='https://www.hntradesp.com/ocrm/tax/salesStatistics/exportExcelData'
    Payload_salesStatistics={
        'invoiceDateStart': "202003",
        'invoiceDateEnd': tomonth,
        }
    head_salesStatistics={'Cookie': cookies_login[0],'Accept-Encoding':''}
    response_salesStatistics=requests.post(url_salesStatistics,data=Payload_salesStatistics,headers=head_salesStatistics,verify=False)
    print(response_salesStatistics.text)
    Fname_salesStatistics=urllib.parse.quote(re.findall('"msg":"(.*?)"', response_salesStatistics.text)[0])
    url_Fname_salesStatistics_real='https://www.hntradesp.com/ocrm/common/download?fileName='+Fname_salesStatistics
    file_info = requests.get(url_Fname_salesStatistics_real,data=Payload_salesStatistics,headers=head_salesStatistics,verify=False)
    with open(r'd:/DATABASE/SYSreporter/'+today+'客户开票.xlsx', 'wb') as file:
        file.write(file_info.content)
    with open(r'd:/DATABASE/all/'+'客户开票.xlsx', 'wb') as file:
        file.write(file_info.content)
    print("开票数据导出完成...")
    
    
    #纳税明细导出
    url_tax=' https://www.hntradesp.com/ocrm/tax/taxStatisticsInfo/asyncExportExcelData'
    Payload_tax={
        'isAsc': 'asc',
        'lateFee': 'T',
        'pageSize': "10",
        'pageNum' : '1'
         }
    head_tax={'Cookie': cookies_login[0],'Accept-Encoding':''}
    response_tax=requests.post(url_tax,data=Payload_tax,headers=head_tax,verify=False)
    print(response_tax.text)
    print('等待纳税导出')
    for e in range(10):
        time.sleep(1)
        print(e)
        
    url_list='https://www.hntradesp.com/ocrm/info/fileInfo/list'
    Payload_list={
        
        'pageSize': "10",
        'pageNum' : '1'
        
        }
    response_list2=requests.post(url_list,headers=head_tax,data=Payload_list,verify=False)    
        
    
    Fname_tax=re.findall('"fileUrl":"(.*?)"', response_list2.text)[0]

    file_tax = requests.get(Fname_tax,headers=head_tax,verify=False)
    with open(r'd:/DATABASE/Tax/'+today+'纳税.xlsx', 'wb') as file:
        file.write(file_tax.content)
    with open(r'd:/DATABASE/all/'+'纳税.xlsx', 'wb') as file:
        file.write(file_tax.content)
    print("纳税数据导出完成...")
    
   
    #销项发票导出
#   遍历型导出所有的销项明细
    url_detail='https://www.hntradesp.com/ocrm/member/invoice/detailexport.ajax'
    df_history=pd.read_csv('d:/DATABASE/Saledetail/21年销项发票明细.csv')
    
    c=(now.year-2022)*12+now.month
    
    for i in range(1,c):
        mo=int((i-0.1)//12)
        temp_date_a=datetime.date(2022+mo, i-mo*12, 1)
        temp_date_b=datetime.date(2022+mo,i-mo*12, calendar.monthrange(2022+mo, i-mo*12)[1])
        file_name=temp_date_b.strftime('%Y-%m')+'销项发票明细.xlsx'
        Payload_detail_recent={
            'startInvoiceTime': temp_date_a.strftime('%Y-%m-%d'),
            'endInvoiceTime': temp_date_b.strftime('%Y-%m-%d'),
            }
        head_detail={'Cookie': cookies_login[0],'Accept-Encoding':''}
        response_detail_recent=requests.post(url_detail,data=Payload_detail_recent,headers=head_detail,verify=False)
        Fname_detail_recent=urllib.parse.quote(re.findall('"msg":"(.*?)"', response_detail_recent.text)[0])
        url_Fname_detail_real='https://www.hntradesp.com/ocrm/common/download?fileName='+Fname_detail_recent
        file_info_detail_recent = requests.get(url_Fname_detail_real,data=Payload_detail_recent,headers=head_detail,verify=False)
        with open(r'd:/DATABASE/Saledetail/'+file_name, 'wb') as file:
            file.write(file_info_detail_recent.content) 
        print(str(2022+mo)+str(i-mo*12)+"销项发票数据导出完成...")
        xlsx_1=pd.ExcelFile('d:/DATABASE/Saledetail/'+file_name)      
        df_1 = pd.read_excel(xlsx_1, 0,dtype={'发票号码':np.str_,'发票代码':np.str_})
        df_history=pd.concat([df_1, df_history],ignore_index=True)  
    df_history.to_csv('d:/DATABASE/all/销项发票明细.csv',index=None,encoding='utf_8_sig')
    df_history.to_csv('d:/DATABASE/Saledetail/'+today+'销项发票明细.csv',index=None,encoding='utf_8_sig')
    print('销项明细已刷新...')
    #返税数据导出
  #  url_tax_back='https://www.hntradesp.com/ocrm/tax/rebate/export.ajax?taxRebateCorpNo=&taxDateStart=&taxDateEnd=&taxRebateCity=&taxTotal='
 #   file_info_tax_back = requests.get(url_tax_back,headers=head_tax,verify=False)
  #  with open(r'd:/DATABASE/Taxback/'+today+'返税.xls', 'wb') as file:
 #       file.write(file_info_tax_back.content)
 #   with open(r'd:/DATABASE/all/'+'返税.xls', 'wb') as file:
 #       file.write(file_info_tax_back.content)
#    print("返税数据导出完成...")
    #会员信息导出
    url_cus='https://www.hntradesp.com/ocrm/tax/customerAccount/exportExcelData?request=customerInfoId=&regCity=&businessAgentId=&referee=&contractSignDateStr=&serviceDateStr=&deptId=&customerInfoId='
    file_info_cus = requests.get(url_cus,headers=head_tax,verify=False)
    with open(r'd:/DATABASE/CustomersInfo/'+today+'会员台账.xls', 'wb') as file:
        file.write(file_info_cus.content)
    with open(r'd:/DATABASE/all/'+'会员台账.xls', 'wb') as file:
        file.write(file_info_cus.content)
    print("会员数据导出完成...")
     #大屏数据下载  
    url_screen='https://www.hntradesp.com/ocrm/tax/customerAccount/toBigScreenDataExport'
    file_tax = requests.get(url_screen,headers=head_tax,verify=False)
    with open(r'd:/DATABASE/screen/'+today+'大屏数据.xls', 'wb') as file:
        file.write(file_tax.content)
    with open(r'd:/DATABASE/all/'+'大屏数据.xls', 'wb') as file:
        file.write(file_tax.content)
    print("大屏数据导出完成...")
    #完税证明导出
    
    url_new_prove='https://www.hntradesp.com/ocrm//tax/taxPaymentAccount/asyncExport.ajax'
    Payload_prove_yp={
        'taxRegCity':'YP'
        
        }
    
    head_prove_yp={'Cookie': cookies_login[0],'Accept-Encoding':''}
    response_prove_yp=requests.post(url_new_prove,data=Payload_prove_yp,headers=head_prove_hk,verify=False)
    print(response_prove_yp.text)
    print('将等待10s')
    time.sleep(10)
    url_list='https://www.hntradesp.com/ocrm/info/fileInfo/list'
    Payload_list={
        
        'pageSize': "10",
        'pageNum' : '1'
        
        }
    response_list=requests.post(url_list,headers=head_prove_yp,data=Payload_list,verify=False)
    
    url_YP_download=re.findall('"fileUrl":"(.*?)"',response_list.text)[0]
    file_info_yp = requests.get(url_YP_download,headers=head_prove_hk,verify=False)
    
    with open(r'd:/DATABASE/Taxproof/'+now.strftime('%Y-%m-%d')+'洋浦完税证明.xlsx', 'wb') as file:
        file.write(file_info_yp.content)
    with open(r'd:/DATABASE/all/'+'洋浦完税证明.xlsx', 'wb') as file:
        file.write(file_info_yp.content)
    print('2021年10月'+'至'+now.strftime('%Y-%m')+"洋浦完税证明导出完成...")   
    
    
  #  response_login=requests.post(url_login,data=Payload_login)
  #  cookies_login =re.findall("'Set-Cookie': '(.*?);", str(response_login.headers))
  #  head_tax={'Cookie': cookies_login[0]}
  #  star='&taxMonthStart='+'202109'
   # end='&taxMonthEnd='+now.strftime('%Y%m')
  #  url_seg='https://www.hntradesp.com/ocrm/tax/taxDutyPaidProof/newExportDetailData?mode=address&taxRegCity=HK&customerNo='+star+end+'&taxInputStartDate=&taxInputEndDate='
 #   file_info = requests.get(url_seg,headers=head_tax)
 #   with open(r'd:/DATABASE/Taxproof/'+now.strftime('%Y-%m')+'海口完税证明.xlsx', 'wb') as file:
  #      file.write(file_info.content)
 #   with open(r'd:/DATABASE/all/'+'海口完税证明.xlsx', 'wb') as file:
 #       file.write(file_info.content)
 #   print('2021年10月'+'至'+now.strftime('%Y-%m')+"海口完税证明导出完成...")   
  #  url_seg='https://www.hntradesp.com/ocrm/tax/taxDutyPaidProof/newExportDetailData?mode=address&taxRegCity=YP&customerNo='+star+end+'&taxInputStartDate=&taxInputEndDate='
 #   file_info = requests.get(url_seg,headers=head_tax)
  #  with open(r'd:/DATABASE/Taxproof/'+now.strftime('%Y-%m')+'洋浦完税证明.xlsx', 'wb') as file:
  #      file.write(file_info.content)
  #  with open(r'd:/DATABASE/all/'+'洋浦完税证明.xlsx', 'wb') as file:
  #      file.write(file_info.content)
  #  print('2021年10月'+'至'+now.strftime('%Y-%m')+"洋浦完税证明导出完成...")    
    print("所有数据导出完成")
def TaxDownload_seg(stardate=202201,enddate=202207):
    url_login='https://www.hntradesp.com/ocrm/login'
    Payload_login={
        'username':'lvchangting',
        'password': 'lvchangting123',
        'rememberMe': 'false'
        }
    response_login=requests.post(url_login,data=Payload_login,verify=False)
    cookies_login =re.findall("'Set-Cookie': '(.*?);", str(response_login.headers))
    head_tax={'Cookie': cookies_login[0]}
    lenth=(int(str(enddate)[:4])-int(str(stardate)[:4]))*12-int(str(stardate)[4:6])+int(str(enddate)[4:6])
    d = datetime.date(int(str(stardate)[:4]),int(str(stardate)[4:6]), 1)
    for i in range(lenth+1):
        this_month_start = datetime.date(d.year+(int(str(stardate)[4:6])+i)//12, d.month+i-(int(str(stardate)[4:6])+i)//13*12, 1)
        this_month_end = datetime.date(d.year+(int(str(stardate)[4:6])+i)//12, d.month+i-(int(str(stardate)[4:6])+i)//13*12, calendar.monthrange(d.year+(int(str(stardate)[4:6])+i)//12, d.month+i-(int(str(stardate)[4:6])+i)//13*12)[1])
        star='taxInputStartDate='+this_month_start.strftime('%Y-%m-%d')
        end='&taxInputEndDate='+this_month_end.strftime('%Y-%m-%d')
        url_seg='https://www.hntradesp.com/ocrm/tax/taxDutyPaidProof/exportDetailData?mode=address&city=HK&customerNo=&taxMonthStart=&taxMonthEnd=&'+star+end
        file_info = requests.get(url_seg,headers=head_tax)
        with open(r'd:/DATABASE/Taxproof/'+this_month_start.strftime('%Y-%m')+'完税证明.xlsx', 'wb') as file:
            file.write(file_info.content)
        print(this_month_start.strftime('%Y-%m')+"完税证明导出完成...")
def TaxDownload(stardate=202201,enddate=202207,place='海口',route=''):
    url_login='https://www.hntradesp.com/ocrm/login'
    Payload_login={
        'username':'lvchangting',
        'password': 'lvchangting123',
        'rememberMe': 'false'
        }
    response_login=requests.post(url_login,data=Payload_login,verify=False)
    cookies_login =re.findall("'Set-Cookie': '(.*?);", str(response_login.headers))
    head_tax={'Cookie': cookies_login[0]}
    d = datetime.date(int(str(stardate)[:4]),int(str(stardate)[4:6]), 1)
    d2=datetime.date(int(str(enddate)[:4]),int(str(enddate)[4:6]), 1)
    this_month_start = datetime.date(d.year, d.month, 1)
    this_month_end = datetime.date(d2.year, d2.month, calendar.monthrange(d2.year, d2.month)[1])
    star='taxInputStartDate='+this_month_start.strftime('%Y-%m-%d')
    end='&taxInputEndDate='+this_month_end.strftime('%Y-%m-%d')
    if palce == '海口':
        url_seg='https://www.hntradesp.com/ocrm/tax/taxDutyPaidProof/newExportDetailData?mode=address&city=HK&customerNo=&taxMonthStart=&taxMonthEnd=&'+star+end
        file_info = requests.get(url_seg,headers=head_tax)
        with open(route+this_month_start.strftime('%Y-%m')+'完税证明.xlsx', 'wb') as file:
            file.write(file_info.content)
        print(this_month_start.strftime('%Y-%m')+'至'+this_month_end.strftime('%Y-%m')+"完税证明导出完成...")        
