# -*- coding: utf-8 -*-
"""
Created on Mon Jul 11 09:40:10 2022

@author: LV
used for mothly cumstomer Reporter
script 
"""



import pandas as pd
import datetime 
import re
from interval import Interval
import warnings
import numpy as np
import pdftest 
#'d:/analyze program/see.xlsx'
warnings.filterwarnings('ignore')

def round1(x,y=2):
    if x>=0:
        return round(x+0.0000001,y)
    else :
        return round(x-0.0000001,y)   

class initialize():
    today=datetime.datetime.now()
    tomonth=int(datetime.datetime.now().strftime('%Y%m'))
    def __init__(self):
        xlsx_1=pd.ExcelFile('d:/DATABASE/all/会员台账.xls')
        xlsx_5 = pd.ExcelFile('c:/Users/LV/Desktop/DATA OF HNIEEC/业务预算/业务预算-招商服务2022-6.xlsx')
        
        self.df_1 = pd.read_excel(xlsx_1, 0)              #df1会员信息台账
        self.df_1['停止合作时间'] = pd.to_datetime(self.df_1['停止合作时间'])
        self.df_1['停止合作时间']=self.df_1['停止合作时间'].dt.strftime('%Y/%m/%d')
        print('member book Data load success')
        self.df_11 = pd.read_excel('d:/DATABASE/all/大屏数据.xls')
        print('screen data load success')
        self.df_2 = pd.read_csv('d:/DATABASE/all/销项发票明细.csv')
        print('Sale detail load success')
        self.df_3 = pd.read_excel('d:/DATABASE/all/客户开票.xlsx')
        
        
                 
        print('value of trade load success')
        self.df_4 = pd.read_excel('d:/DATABASE/all/纳税.xlsx')
        self.df_41=pd.read_excel('d:/DATABASE/all/返税.xls',header=1)
        print('Tax Data load success')
        self.df_5 = pd.read_excel(xlsx_5, "收入明细表",header=1)
        print('financial statement-income load success')
        self.df_6 = pd.read_excel(xlsx_5, 1)                           #df6
        print('financial statement-others load success')
        self.df_7 = pd.read_excel('d:/DATABASE/taxrate/返税比例.xlsx') #df7返税比例
        self.df_8 = pd.read_excel('d:/DATABASE/all/洋浦完税证明.xlsx',header=1)  #df8完税证明
        self.df_81 = pd.read_excel('d:/DATABASE/all/海口完税证明.xlsx',header=1)  #df8完税证明
        self.df_9 = pd.read_excel('d:/DATABASE/cargotype/产品字典库.xlsx')  #df9产品字典库
        self.df_10 = pd.read_excel('d:/DATABASE/Customertype/customertype.xlsx',dtype={'编码':str})  #df10 洋浦企业类型字典
        self.df_yp_applyed =  pd.read_excel('d:/DATABASE/all/ypapplyed.xlsx') #df_yp_applyed 洋浦已申请清单
        
       
        self.df_10['官方类型'] = np.where(self.df_10['官方类型']=='能源','能源类',self.df_10['官方类型'])
        self.df_10['官方类型'] = np.where(self.df_10['官方类型']=='非能源','非能源类',self.df_10['官方类型'])
        self.df_11 = pd.read_excel('d:/DATABASE/Customertype/custype.xlsx')
        
        self.df_tax_real = pd.read_excel('d:/DATABASE/Customertype/custype.xlsx')
        
        
        self.df_applyed = pd.read_excel('d:/DATABASE/all/applyed.xlsx') #df_applyed 已申请完税证明字典
        self.df_charge_rate = pd.read_excel('d:/DATABASE/service_charge/rate.xlsx') #df_charge_rate 返税服务费比例字典
        self.list_re_cargo=list(self.df_9['key'])
        self.dict_type=dict(zip(self.df_9['key'],self.df_9['type']))
        print('All Data load success')
        self.namedict = dict(zip(self.df_1['统一社会信用代码'],self.df_1['企业名称']))

        self.df_2['开票月份']=self.df_2['开票日期'].str[:7]
        self.df_3['销方企业名称']=self.df_3['统一社会信用代码'].apply(lambda x :self.namedict.get(x))     
    
    def add_edu(self,df,cor=29):
        
        pass
    def taxprove_remake(self):
        pass
        
    def yp_rete_get(self,ID,taxtype):

        cs_type=dict(zip(self.df_11['统一社会信用代码'],self.df_11['类型']))
        if cs_type.get(ID) == '非能源类':
            
            return self.df_7[self.df_7['类型']=='非能源类'][taxtype]
        else:
            return self.df_7[self.df_7['类型']=='能源类'][taxtype]
            
            
            
            
            
        
    def tax_remake(self,place='海口'):
        raw=pdftest.run(place=place)
        raw['金额']=raw['金额'].astype('float')
        raw['金额']=raw['金额'].apply(lambda x: format(x,'.2f'))
        df_80=self.df_8.append(self.df_81,ignore_index=True)
        
        df_80.fillna(0,inplace=True)
        df_80['纳税合计']=df_80['增值税']+df_80['城建税']+df_80['教育费附加']+df_80['地方教育附加']+df_80['印花税']+df_80['企业所得税']+df_80['个人所得税']+df_80['其他收入-工会经费']
        df_80['纳税合计']=df_80['纳税合计'].apply(lambda x: format(x,'.2f'))
        df_80['mark']=df_80['Unnamed: 5']+df_80['起始日']+df_80['截止日']+df_80['Unnamed: 14']+df_80['纳税合计'].astype('str')
        raw['起始日']= pd.to_datetime(raw['起始日'])
        raw['截止日']= pd.to_datetime(raw['截止日'])
        raw['税款入库日期']= pd.to_datetime(raw['税款入库日期'])
        raw['起始日']= raw['起始日'].apply(lambda x: x.strftime('%Y/%m/%d'))
        raw['截止日']= raw['截止日'].apply(lambda x: x.strftime('%Y/%m/%d'))
        raw['税款入库日期']= raw['税款入库日期'].apply(lambda x: x.strftime('%Y/%m/%d'))
        raw['mark']=raw['统一社会信用代码']+raw['起始日']+raw['截止日']+raw['税种']+raw['金额']
        
        IDlist=list(raw['统一社会信用代码'])
        
        marklist=list(raw['mark'])
        
        
        cus_ID_dict=dict(zip(df_80['Unnamed: 5'],df_80['Unnamed: 0']))
        tax_ID_dict=dict(zip(df_80['mark'],df_80['Unnamed: 2']))
        tax_sysID_dict=dict(zip(df_80['mark'],df_80['Unnamed: 1']))
        tax_place_dict=dict(zip(self.df_1['统一社会信用代码'],self.df_1['注册地']))
        cus_singed_dict=dict(zip(df_80['Unnamed: 5'],df_80['Unnamed: 7']))
        cus_sign_date=dict(zip(self.df_1['统一社会信用代码'],self.df_1['签约日期']))
        cus_sign_date_end=dict(zip(self.df_1['统一社会信用代码'],self.df_1['服务期限终止']))
        cus_reg_date=dict(zip(self.df_1['统一社会信用代码'],self.df_1['成立(迁移)日期']))
        zhina_dict=dict(zip(df_80['mark'],df_80['是否为滞纳金']))
        
        
        cus_ID_list=[]
        tax_ID_list=[]
        tax_sysID_list=[]
        tax_place_list=[]
        cus_singed_list=[]
        cus_sign_date_list=[]
        cus_sign_date_end_list=[]
        cus_reg_date_list=[]
        namelist=[]
        zhina=[]
        tax_person=[]
        
        
        
        for i in range(len(IDlist)):
            namelist.append(self.namedict.get(IDlist[i]))
            cus_ID_list.append(cus_ID_dict.get(IDlist[i]))
            tax_ID_list.append(tax_ID_dict.get(marklist[i]))
            tax_sysID_list.append(tax_sysID_dict.get(marklist[i]))
            tax_place_list.append(tax_place_dict.get(IDlist[i]))
            cus_singed_list.append(cus_singed_dict.get(IDlist[i]))
            cus_sign_date_list.append(cus_sign_date.get(IDlist[i]))
            cus_sign_date_end_list.append(cus_sign_date_end.get(IDlist[i]))
            cus_reg_date_list.append(cus_reg_date.get(IDlist[i]))
            zhina.append(zhina_dict.get(marklist[i]))
        raw.insert(0,'Unnamed: 0',cus_ID_list)
        raw.insert(1,'Unnamed: 1',tax_sysID_list)
        raw.insert(2,'Unnamed: 2',tax_ID_list)
        raw.insert(4,'Unnamed: 4',tax_place_list)
        raw.insert(6,'Unnamed: 6',namelist)
        raw.insert(7,'Unnamed: 7',cus_singed_list)
        raw.insert(8,'Unnamed: 8',cus_sign_date_list)
        raw.insert(9,'Unnamed: 9',cus_sign_date_end_list)
        raw.insert(10,'Unnamed: 10',cus_reg_date_list)
        tax_type=list(raw['税种'])
        raw.drop(columns=['税种','入库地'],inplace=True)
        raw.insert(14,'Unnamed: 14',tax_type)
        tax_date=list(raw['截止日'])
        for b in range(len(tax_date)):
            tax_date[b]=int(tax_date[b][:4]+tax_date[b][5:7])
        raw.insert(15,'Unnamed: 15',tax_date) 
        
        money=list(raw['金额'])
        zz=[0 for i in range(len(raw))]
        cj=[0 for i in range(len(raw))]
        jy=[0 for i in range(len(raw))]
        jy_df=[0 for i in range(len(raw))]
        yh=[0 for i in range(len(raw))]
        qy=[0 for i in range(len(raw))]
        gr=[0 for i in range(len(raw))]
        qt=[0 for i in range(len(raw))]
        qt=[0 for i in range(len(raw))]
        apply_date=['' for i in range(len(raw))]
        count=0
        for c in tax_type:
            if c== '企业所得税' :
                qy[count]=float(money[count])
            if c== '印花税' :
                yh[count]=float(money[count])
            if c== '地方教育附加' :
                jy_df[count]=float(money[count])
            if c== '城市维护建设税' :
                cj[count]=float(money[count])
            if c== '增值税' :
                zz[count]=float(money[count])
            if c== '教育费附加' :
                jy[count]=float(money[count])
            if c== '个人所得税' :
                gr[count]=float(money[count])
            count+=1
        raw.insert(16,'增值税',zz)
        raw.insert(17,'城建税',cj)
        raw.insert(18,'教育费附加',jy)
        raw.insert(19,'地方教育附加',jy_df)
        raw.insert(20,'印花税',yh)
        raw.insert(21,'企业所得税',qy)  
        raw.insert(22,'个人所得税',gr)         
        raw.insert(23,'其他收入-工会经费',qt) 
        raw.drop(columns=['金额'],inplace=True)
        raw.insert(24,'是否为滞纳金',zhina) 
        df_80['mark2']=df_80['Unnamed: 5']+df_80['Unnamed: 15'].astype('str')
        tax_person_dict=dict(zip(df_80['mark2'],df_80['Unnamed: 25']))  
        raw['mark2']=raw['统一社会信用代码']+raw['Unnamed: 15'].astype('str')
        mark2list=list(raw['mark2'])
        
        for e in range(len(mark2list)):
            tax_person.append(tax_person_dict.get(mark2list[e]))
        raw.insert(25,'Unnamed: 25',tax_person) 
        raw.insert(26,'Unnamed: 26',apply_date) 
      
        if place == '海口':
            raw['增值税.1'] = np.where(raw['Unnamed: 4']!='',raw['增值税']*0.425,0 )
            raw['增值税.1']=raw['增值税.1'].apply(lambda x :round1(x,2))
            raw['城建税.1'] = np.where(raw['Unnamed: 4']!='',raw['城建税']*0.85, 0 )
            raw['城建税.1']=raw['城建税.1'].apply(lambda x :round1(x,2))
            raw['印花税.1'] = np.where(raw['Unnamed: 4']!='',raw['印花税']*0.85,0 )
            raw['印花税.1']=raw['印花税.1'].apply(lambda x :round1(x,2))
            raw['企业所得税.1'] = np.where(raw['Unnamed: 4']!='',raw['企业所得税']*0.34, 0 )
            raw['企业所得税.1']=raw['企业所得税.1'].apply(lambda x :round1(x,2))
            raw['个人所得税.1'] = np.where(raw['Unnamed: 4']!='',raw['个人所得税']*0.4, 0 )
            raw['个人所得税.1']=raw['个人所得税.1'].apply(lambda x :round1(x,2))
        else:
            raw['增值税.1']=0
            raw['城建税.1']=0
            raw['印花税.1']=0
            raw['企业所得税.1']=0
            raw['个人所得税.1']=0
            for r in range(len(raw)):
                raw.iloc[r,29]=raw.loc[r,'增值税']*self.yp_rete_get(raw.loc[r,'统一社会信用代码'],'增值税')
                raw.iloc[r,30]=raw.loc[r,'城建税']*self.yp_rete_get(raw.loc[r,'统一社会信用代码'],'城建税')
                raw.iloc[r,31]=raw.loc[r,'印花税']*self.yp_rete_get(raw.loc[r,'统一社会信用代码'],'印花税')
                raw.iloc[r,32]=raw.loc[r,'企业所得税']*self.yp_rete_get(raw.loc[r,'统一社会信用代码'],'企业所得税')
                raw.iloc[r,33]=raw.loc[r,'个人所得税']*self.yp_rete_get(raw.loc[r,'统一社会信用代码'],'个人所得税')
        
        raw['Unnamed: 32'] = raw['增值税.1']+raw['城建税.1']+raw['印花税.1']+raw['企业所得税.1']+raw['个人所得税.1']

        #合计平衡原则
        raw['合计']=raw['增值税.1']+raw['城建税.1']+raw['印花税.1']+raw['企业所得税.1']
        raw['省级财力贡献']=raw['增值税.1']*0.3+raw['城建税.1']*0.2+raw['企业所得税.1']*0.3
        raw['省级财力贡献']=raw['省级财力贡献'].apply(lambda x :round1(x,2))
        raw['市级财力贡献']=raw['增值税.1']*0.7+raw['城建税.1']*0.8+raw['印花税.1']+raw['企业所得税.1']*0.7
        raw['市级财力贡献']=raw['市级财力贡献'].apply(lambda x :round1(x,2))
        raw['合计']=raw['省级财力贡献']+raw['市级财力贡献']
        raw['Unnamed: 36']=''
        raw['Unnamed: 37']=''
        raw['Unnamed: 38']=''
        raw['Unnamed: 39']=''
        raw['Unnamed: 40']=''
        raw['Unnamed: 41']=''
        raw['Unnamed: 42']=''
        raw['Unnamed: 43']=''
        raw['Unnamed: 44']=''
        raw['Unnamed: 45']=''
        raw.drop(columns=['mark','mark2'],inplace=True)
        raw['Unnamed: 8']= pd.to_datetime(raw['Unnamed: 8'])
        raw['Unnamed: 9']= pd.to_datetime(raw['Unnamed: 9'])
        raw['Unnamed: 10']= pd.to_datetime(raw['Unnamed: 10'])
        raw['Unnamed: 10']= raw['Unnamed: 10'].apply(lambda x: x.strftime('%Y/%m/%d'))
        np.where(self.df_10['官方类型']=='能源','能源类',self.df_10['官方类型'])
        
        li8=list(raw['Unnamed: 8'])
        li9=list(raw['Unnamed: 9'])

        for d in range(len(li8)):
            try:
                li8[d]=li8[d].strftime('%Y/%m/%d')
                li9[d]=li9[d].strftime('%Y/%m/%d')
            except:
                pass
                
        
        raw['Unnamed: 8']= li8
        raw['Unnamed: 9']= li9
        
        raw.rename(columns={"统一社会信用代码":'Unnamed: 5',"完税证明编号":'Unnamed: 3',"税款入库日期":'Unnamed: 13'},inplace=True)
        raw['Unnamed: 8'].fillna('',inplace=True)
        #raw=raw[raw['截止日']>=raw['Unnamed: 8'] ]
        self.df_8old=self.df_8
        self.df_81old=self.df_81
        
        
        
        
        if place == '海口':
            self.df_81=raw
            mask=(self.df_81old['Unnamed: 13'] >= '2022/01/01') & (self.df_81old['Unnamed: 13'] <= '2022/10/31')
            df_82old=self.df_81old.loc[mask]
            new_cuslist=list(set(self.df_81['Unnamed: 5']))
        else :
            self.df_8=raw
           
            df_82old=self.df_8old
            new_cuslist=list(set(self.df_8['Unnamed: 5']))
        df_82old=df_82old[df_82old['Unnamed: 14']!='个人所得税']
        df_82old=df_82old[df_82old['Unnamed: 14']!='其他收入']
        df_82old=df_82old[df_82old['Unnamed: 14']!='城镇土地使用税']
        df_82old=df_82old[df_82old['Unnamed: 14']!='基本医疗保险费']
        df_82old=df_82old[df_82old['Unnamed: 14']!='失业保险费']
        df_82old=df_82old[df_82old['Unnamed: 14']!='工伤保险费']
        df_82old=df_82old[df_82old['Unnamed: 14']!='房产税']
        df_82old=df_82old[df_82old['Unnamed: 14']!='职工基本养老保险(个险费']
        
        
        old_cuslist=list(set(df_82old['Unnamed: 5']))
        
        # 未导出的企业名单
        needask=set(old_cuslist).difference(set(new_cuslist))
        needaskname=[]
        for t in needask:
            needaskname.append(self.namedict.get(t))
            
        need_df=pd.DataFrame()
        need_df['企业名称']=needaskname
        need_df['统一社会信用代码']=list(needask)
        need_df=self.add_end_date(need_df, 1)
        need_df=self.add_tax_person(need_df,'统一社会信用代码','统一社会信用代码')
        filename='d:/analyze program/'+place+'未提交的公司.xlsx'
        need_df.to_excel(filename)
        # 多导出的的企业名单
        needinmport=set(new_cuslist).difference(set(old_cuslist))
        return raw
        
        
        
        
        
        
    def get_last_renew(self):
        self.last_renew_df=self.df_1.loc[:,['统一社会信用代码','第一年','第二年','第三年']]
        liID=list(self.last_renew_df['统一社会信用代码'])
        li1=list(self.last_renew_df['第一年'])
        li2=list(self.last_renew_df['第二年'])
        li3=list(self.last_renew_df['第三年'])
        li4=[]
        for x in range(len(li1)):
            if li3[x] != li3[x]:
                if li2[x] != li2[x]:
                    if li1[x] != li1[x]:
                        pass
                    else:
                        last_re=li1[x]
                else:
                    last_re=li2[x]
                    if li2[x]<li1[x]:
                        print(self.namedict.get(liID[x])+'第二年续费时间小于第一年')
                
            else:
                if li3[x]<li2[x]:
                    print(self.namedict.get(liID[x])+'第三年续费时间小于第二年')
                last_re=li3[x]
            li4.append(last_re)
            
        self.last_renew_df['最后续费时间'] = li4
        self.last_renew_dict = dict(zip(self.last_renew_df['统一社会信用代码'],self.last_renew_df['最后续费时间']))
        
    def add_last_renew(self,raw_pd,a,p=0):
        self.get_last_renew()
        listname=[]
        if a==0:
            for i in raw_pd.index:
                listname.append(self.last_renew_dict.get(i))
            raw_pd.insert(p,'最后续费时间',listname)
        else:
            for i in list(raw_pd['统一社会信用代码']):
                listname.append(self.last_renew_dict.get(i))
            raw_pd.insert(p,'最后续费时间',listname)
        return raw_pd
    
    def add_sign_date(self,raw_pd,a,p=0):
        self.sign_dict = dict(zip(self.df_1['统一社会信用代码'],self.df_1['签约日期']))
        listname=[]
        if a==0:
            for i in raw_pd.index:
                listname.append(self.sign_dict.get(i))
            raw_pd.insert(p,'签约时间',listname)
        else:
            for i in list(raw_pd['统一社会信用代码']):
                listname.append(self.sign_dict.get(i))
            raw_pd.insert(p,'签约时间',listname)
        return raw_pd
    
    def add_reg_date(self,raw_pd,a,p=0):
        self.sign_dict = dict(zip(self.df_1['统一社会信用代码'],self.df_1['成立(迁移)日期']))
        listname=[]
        if a==0:
            for i in raw_pd.index:
                listname.append(self.sign_dict.get(i))
            raw_pd.insert(p,'成立(迁移)日期',listname)
        else:
            for i in list(raw_pd['统一社会信用代码']):
                listname.append(self.sign_dict.get(i))
            raw_pd.insert(p,'成立(迁移)日期',listname)
        return raw_pd
    def add_serve_start(self,raw_pd,a,p=0):
        self.sign_dict = dict(zip(self.df_1['统一社会信用代码'],self.df_1['服务期限起始']))
        listname=[]
        if a==0:
            for i in raw_pd.index:
                listname.append(self.sign_dict.get(i))
            raw_pd.insert(p,'服务期限起始',listname)
        else:
            for i in list(raw_pd['统一社会信用代码']):
                listname.append(self.sign_dict.get(i))
            raw_pd.insert(p,'服务期限起始',listname)
        return raw_pd
        
        
   #预缴管控
    def yp_needpay(self,start=202201,end=tomonth):
        raw_data=self.service_charge(from_date=start,end_date=end)
        mask=(raw_data['税收所属期'] >= start) & (raw_data['税收所属期'] <= end)
        userange=raw_data.loc[mask]
        yp_userange=userange[userange['纳税地']=='洋浦']
        yp_result=yp_userange.groupby(['统一社会信用代码']).sum()
        list_yp_apply=yp_result['申请合计']
        list_yp_needpay=[]
        list_yp_tag=[]
        df_10=self.df_10
        df_10.fillna(1,inplace=True)
        name_df=df_10[df_10['官方类型']!=1]
        name_df=name_df[name_df['编码']!=1]

        #目前按照条件满足的做为名单，后续需要调整
        applyedlist=list(name_df['统一社会信用代码'])
        for i in list_yp_apply:
            a=0
            if i>=100000:
                a=i*0.9
            else:
                if i>=50000:
                    a=i*0.8
                else:
                    if i>=10000:
                        a=i*0.7
                    else:
                        a=0
            list_yp_needpay.append(a)
        yp_result['预缴金额']= list_yp_needpay
        yp_result['税收所属期']=str(start)+'--'+str(end)
        
        for x in yp_result.index:
            if x in applyedlist:
                list_yp_tag.append('已申请')
            else:
                list_yp_tag.append('未申请')
                
                
            
        yp_result=self.add_name(yp_result, 0)
        yp_result['是否已申请']=list_yp_tag
        

        
        
        
        return yp_result
            
        
        
        
        
        
   # 新的返税服务费计算方式
    def service_charge(self,from_date=202201,end_date=tomonth):
        self.tax_refresh()
        print('应返补正完成')
        pay_dict=dict(zip(self.df_charge_rate['统一社会信用代码'],self.df_charge_rate['收费模式']))
        payrate_dict=dict(zip(self.df_charge_rate['统一社会信用代码'],self.df_charge_rate['收费比例']))
        mask=(self.df_41['税收所属期'] >= from_date) & (self.df_41['税收所属期'] <= end_date)
        userange=self.df_41.loc[mask]
        for i in range(len(userange)):
            Serie=userange.iloc[i]
            ID=Serie.统一社会信用代码
            date=Serie['税收所属期']
            apply_sum=Serie['申请合计']
            real_sum=Serie['实返合计']
            index_num=self.df_41[(self.df_41.统一社会信用代码==ID)&(self.df_41.税收所属期==date)].index.tolist()
            
            if pay_dict.get(ID) == 1:
                income_sum=apply_sum * payrate_dict.get(ID)
                income_real_sum=real_sum * payrate_dict.get(ID)            
                self.df_41.loc[index_num,'预计收入']= round(income_sum,2)
                self.df_41.loc[index_num,'实返收入']= round(income_real_sum,2)
                print('1被执行')
            else:
                if pay_dict.get(ID) == 2:               
                    a=self.period_sale(date, ID)
                    try:
                        b=float(self.df_3[(self.df_3.统一社会信用代码==ID)&(self.df_3.开票月份==date)].价税合计)
                    except:
                        b=0
                        print(ID+str(date)+'价税合计获取错误')
                    if (a+b)<=500000000:
                        n1=b
                        n2=0
                        n3=0
                    else:
                        if (a+b) <=1000000000:
                            n3=0
                            n2=a+b-500000000
                            if a <=500000000:
                                n1=500000000-a
                            else:
                                n1=0
                        else:
                            n3=(a+b)-1000000000
                            if a >=1000000000:
                                n1=0
                                n2=0
                            else:
                                n2=1000000000-a
                                if a >=500000000:
                                    n1=0
                                else:
                                    n1=500000000-a
                    if n1+n2+n3==0:
                        rate=pd.Series([1,0,0])
                    else:
                        rate=pd.Series([n1/(n1+n2+n3),n2/(n1+n2+n3),n3/(n1+n2+n3)])
                    rate2=pd.Series([0.1,0.05,0.03])
                    income_sum=sum(apply_sum * rate * rate2)
                    income_real_sum=sum(real_sum * rate * rate2)
                    index_num=self.df_41[(self.df_41.统一社会信用代码==ID)&(self.df_41.税收所属期==date)].index.tolist()
                    self.df_41.loc[index_num,'预计收入']= round(income_sum,2)
                    self.df_41.loc[index_num,'实返收入']= round(income_real_sum,2)
                    self.df_41.loc[index_num,'服务期内开票额']=a+b
                    print('2被执行')
                else:
                    self.df_41.loc[index_num,'预计收入']= round(apply_sum *0.1,2)
                    self.df_41.loc[index_num,'实返收入']= round(real_sum*0.1,2)
                    print('3被执行')
                   
        return self.df_41     
    def income_raw_cube(self,DF='',DF_on='',data_name='公司名称',Date_tag='所属月份',apply_tag='申请合计'):
        

        pay_dict=dict(zip(self.df_charge_rate['统一社会信用代码'],self.df_charge_rate['收费模式']))
        payrate_dict=dict(zip(self.df_charge_rate['统一社会信用代码'],self.df_charge_rate['收费比例']))
        sale_dict=dict(zip(self.df_3['统一社会信用代码']+self.df_3['开票月份'].astype('str'),self.df_3['价税合计']))
        
        userange=DF

        sign_dict = dict(zip(self.df_1['统一社会信用代码'],self.df_1['服务期限起始']))
        IDdict = dict(zip(self.df_1['企业名称'],self.df_1['统一社会信用代码']))
        
        
        
        for i in range(len(userange)):
            Serie=userange.iloc[i]
            name=Serie[data_name]
            date=int(Serie[Date_tag])
            real_sum=Serie[apply_tag]
            
            try:
                ID=IDdict[name]
            except:
                print( name+'找不到统一信用代码')
            sale_key=ID+str(date)
            
            
            if pay_dict.get(ID) == 1:
                
                income_real_sum=real_sum * payrate_dict.get(ID)            
                userange.loc[i,'实返收入']= round(income_real_sum,2)
                userange.loc[i,'服务期内开票额']=''
                userange.loc[i,'服务费比例']= '{:.0%}'.format(payrate_dict.get(ID))
                userange.loc[i,'服务起始日期']= sign_dict.get(ID) 
                userange.loc[i,'价税合计']= sale_dict.get(sale_key) 
                #print('1被执行')
            else:
                if pay_dict.get(ID) == 2:     
                    detal_df=pd.DataFrame()
                    
                    
                    
                    
                    a=self.period_sale(date, ID)
                    try:
                        b=float(self.df_3[(self.df_3.统一社会信用代码==ID)&(self.df_3.开票月份==date)].价税合计)
                    except:
                        b=0
                        print(ID+str(date)+'价税合计获取错误')
                    if (a+b)<=500000000:
                        n1=b
                        n2=0
                        n3=0
                    else:
                        if (a+b) <=1000000000:
                            n3=0
                            n2=a+b-500000000
                            if a <=500000000:
                                n1=500000000-a
                            else:
                                n1=0
                        else:
                            n3=(a+b)-1000000000
                            if a >=1000000000:
                                n1=0
                                n2=0
                            else:
                                n2=1000000000-a
                                if a >=500000000:
                                    n1=0
                                else:
                                    n1=500000000-a
                    if n1+n2+n3==0:
                        rate=pd.Series([1,0,0])
                    else:
                        rate=pd.Series([n1/(n1+n2+n3),n2/(n1+n2+n3),n3/(n1+n2+n3)])
                    rate2=pd.Series([0.1,0.05,0.03])
                    
                    income_real_sum=sum(real_sum * rate * rate2)
                    

                    userange.loc[i,'实返收入']= round(income_real_sum,2)
                    userange.loc[i,'服务期内开票额']=a+b
                    
                    if rate[0] ==1:
                        userange.loc[i,'服务费比例']= "10%"
                    else:
                        if rate[0] != 0:
                            userange.loc[i,'服务费比例']= "10%-5%"
                        else:
                            if rate[1] != 1 :
                                if rate[1] != 0:
                                    userange.loc[i,'服务费比例']= "5%-3%"
                                else:
                                    userange.loc[i,'服务费比例']= "3%"
                            else:
                                
                                userange.loc[i,'服务费比例']= "5%"
                                
                    
                    userange.loc[i,'服务起始日期']= sign_dict.get(ID) 
                    userange.loc[i,'价税合计']= sale_dict.get(sale_key) 
                   # print('2被执行')
                else:

                    userange.loc[i,'实返收入']= round(real_sum*0.1,2)
                    userange.loc[i,'服务期内开票额']=""
                    userange.loc[i,'服务费比例']= "10%"
                    userange.loc[i,'服务起始日期']= sign_dict.get(ID) 
                    userange.loc[i,'价税合计']= sale_dict.get(sale_key) 
                    
                    
                    
                 #   print('3被执行')
        userange.to_excel('d:/analyze program/海口服务费收入5.xlsx',index=False)          
        return userange   
        
        
    def income_seg(self,DF='',DF_on='',Data_on='',Date_tag='Unnamed: 2',apply_tag='申请合计',place='洋浦'):
        
        DF=pd.read_excel('d:/analyze program/0112需要计算的服务费.xlsx',header=1)
        pay_dict=dict(zip(self.df_charge_rate['统一社会信用代码'],self.df_charge_rate['收费模式']))
        payrate_dict=dict(zip(self.df_charge_rate['统一社会信用代码'],self.df_charge_rate['收费比例']))
        sale_dict=dict(zip(self.df_3['统一社会信用代码']+self.df_3['开票月份'].astype('str'),self.df_3['价税合计']))
        userange=DF
        if place== '洋浦':
            userange=userange[userange[Date_tag]!='留抵']
            userange[Date_tag]= pd.to_datetime(userange[Date_tag])
            userange[Date_tag]= userange[Date_tag].apply(lambda x: x.strftime('%Y%m'))
            
        else:
            userange=userange[userange[Date_tag]!='留抵']

        sign_dict = dict(zip(self.df_1['统一社会信用代码'],self.df_1['服务期限起始']))
        IDdict = dict(zip(self.df_1['企业名称'],self.df_1['统一社会信用代码']))
        
        
        for i in range(len(userange)):
            Serie=userange.iloc[i]
            name=Serie['Unnamed: 1']
            date=int(Serie[Date_tag])
            real_sum=Serie[apply_tag]
            
            try:
                ID=IDdict[name]
            except:
                print( name+'找不到统一信用代码')
            sale_key=ID+str(date)
            
            
            if pay_dict.get(ID) == 1:
                
                income_real_sum=real_sum * payrate_dict.get(ID)            
                userange.loc[i,'实返收入']= round(income_real_sum,2)
                userange.loc[i,'服务期内开票额']=''
                userange.loc[i,'服务费比例']= '{:.0%}'.format(payrate_dict.get(ID))
                userange.loc[i,'服务起始日期']= sign_dict.get(ID) 
                userange.loc[i,'价税合计']= sale_dict.get(sale_key) 
                #print('1被执行')
            else:
                if pay_dict.get(ID) == 2:     
                    detal_df=pd.DataFrame()
                    
                    
                    
                    
                    a=self.period_sale(date, ID)
                    try:
                        b=float(self.df_3[(self.df_3.统一社会信用代码==ID)&(self.df_3.开票月份==date)].价税合计)
                    except:
                        b=0
                        print(ID+str(date)+'价税合计获取错误')
                    if (a+b)<=500000000:
                        n1=b
                        n2=0
                        n3=0
                    else:
                        if (a+b) <=1000000000:
                            n3=0
                            n2=a+b-500000000
                            if a <=500000000:
                                n1=500000000-a
                            else:
                                n1=0
                        else:
                            n3=(a+b)-1000000000
                            if a >=1000000000:
                                n1=0
                                n2=0
                            else:
                                n2=1000000000-a
                                if a >=500000000:
                                    n1=0
                                else:
                                    n1=500000000-a
                    if n1+n2+n3==0:
                        rate=pd.Series([1,0,0])
                    else:
                        rate=pd.Series([n1/(n1+n2+n3),n2/(n1+n2+n3),n3/(n1+n2+n3)])
                    rate2=pd.Series([0.1,0.05,0.03])
                    
                    income_real_sum=sum(real_sum * rate * rate2)
                    

                    userange.loc[i,'实返收入']= round(income_real_sum,2)
                    userange.loc[i,'服务期内开票额']=a+b
                    
                    if rate[0] ==1:
                        userange.loc[i,'服务费比例']= "10%"
                    else:
                        if rate[0] != 0:
                            userange.loc[i,'服务费比例']= "10%-5%"
                        else:
                            if rate[1] != 1 :
                                if rate[1] != 0:
                                    userange.loc[i,'服务费比例']= "5%-3%"
                                else:
                                    userange.loc[i,'服务费比例']= "3%"
                            else:
                                
                                userange.loc[i,'服务费比例']= "5%"
                                
                    
                    userange.loc[i,'服务起始日期']= sign_dict.get(ID) 
                    userange.loc[i,'价税合计']= sale_dict.get(sale_key) 
                   # print('2被执行')
                else:

                    userange.loc[i,'实返收入']= round(real_sum*0.1,2)
                    userange.loc[i,'服务期内开票额']=""
                    userange.loc[i,'服务费比例']= "10%"
                    userange.loc[i,'服务起始日期']= sign_dict.get(ID) 
                    userange.loc[i,'价税合计']= sale_dict.get(sale_key) 
                    
                    
                    
                 #   print('3被执行')
        userange.to_excel('d:/analyze program/0112服务费预计算.xlsx',index=False)          
        return userange   
    
    
    #税款刷新drop参数设置是否丢弃不满足贸易额条件的税款
    def tax_refresh(self,start=202201,end=tomonth,place='洋浦',drop=False):
        if drop == False:
            result=self.yp_apply(start=start,end=end,method='all',need='raw')
        if drop == True:
            result=self.yp_apply(start=start,end=end,method='all',need='result')
        result=result[result['注册地']!='海口']

        for i in range(len(result)):
            Serie=result.iloc[i]
            ID=Serie['统一社会信用代码_x']
            date=Serie['税收所属期']
            index_num=self.df_41[(self.df_41.统一社会信用代码==ID)&(self.df_41.税收所属期==date)].index.tolist()
            if index_num== []:
                print(ID+str(date))
            self.df_41.loc[index_num,'申请合计']= Serie['合计_x']
            self.df_41.loc[index_num,'增值税.1']= Serie['增值税.1_x']
            self.df_41.loc[index_num,'城建税.1']= Serie['城建税.1_x']
            self.df_41.loc[index_num,'教育费附加.1']= Serie['教育费附加.1_x']
            self.df_41.loc[index_num,'印花税.1']= Serie['印花税.1_x']
            self.df_41.loc[index_num,'企业所得税.1']= Serie['企业所得税.1_x']
            #self.df_41.loc[index_num,'个税.1']= Serie['个人所得税.1_x']
       #needregedit参数用来指定是否按照已完成认定的企业进行申请。method 参数用来确定是否剔除已申请的税款， monthlimit参数用来确定按月还是累计来确定是否剔除。
    def yp_apply(self,needregedit=False,start=202201,end=tomonth,place='全部',method='all',monthlimit=False,salestart=202201,saleend=tomonth,confirm=False,need='raw'):
        
            
        yp_df=self.df_8.iloc[:,:36]
        df_applyed=self.df_applyed
        listID=list(df_applyed['ID'])
        filter_condition={'Unnamed: 1':listID}
        mask=(yp_df['Unnamed: 15'] >= start) & (yp_df['Unnamed: 15'] <= end)
        userange=yp_df.loc[mask]

        needapply=userange[~userange.isin(filter_condition)['Unnamed: 1']]
        applyed=userange[userange.isin(filter_condition)['Unnamed: 1']]
        
        tax=userange.groupby(['Unnamed: 5','Unnamed: 15']).sum()
        tax_unapply=needapply.groupby(['Unnamed: 5','Unnamed: 15']).sum()
        tax_apply=applyed.groupby(['Unnamed: 5','Unnamed: 15']).sum()
        
        tax.reset_index(inplace=True)
        tax.rename(columns={'Unnamed: 5':'统一社会信用代码','Unnamed: 15':'税收所属期'},inplace=True)
        tax=self.add_name(tax,1)
        tax.rename(columns={'统一社会信用代码':'统一社会信用代码_x'},inplace=True)
                            
        tax_unapply.reset_index(inplace=True)
        tax_unapply.rename(columns={'Unnamed: 5':'统一社会信用代码','Unnamed: 15':'税收所属期'},inplace=True)
        tax_unapply=self.add_name(tax_unapply,1)
        tax_unapply.rename(columns={'统一社会信用代码':'统一社会信用代码_y'},inplace=True)
        
        tax_apply.reset_index(inplace=True)
        tax_apply.rename(columns={'Unnamed: 5':'统一社会信用代码','Unnamed: 15':'税收所属期'},inplace=True)
        tax_apply=self.add_name(tax_apply,1)
        tax_apply.rename(columns={'统一社会信用代码':'统一社会信用代码_z'},inplace=True)
        
        tax_1=pd.merge(tax,tax_apply,on=['企业名称','税收所属期'],how='left')
        tax_all=pd.merge(tax_1,tax_unapply,on=['企业名称','税收所属期'],how='left')
        
        
        
        mask_sale=(self.df_3['开票月份'] >= salestart) & (self.df_3['开票月份'] <= saleend) 
        df_3=self.df_3.loc[mask_sale]
        rawdata=pd.merge(tax_all,df_3,left_on=['企业名称','税收所属期'],right_on=['销方企业名称','开票月份'],how='outer')
        rawdata['企业名称'] = np.where(rawdata['企业名称']!=rawdata['企业名称'],rawdata['销方企业名称'],rawdata['企业名称'])
        rawdata['统一社会信用代码_x'] = np.where(rawdata['统一社会信用代码_x']!=rawdata['统一社会信用代码_x'],rawdata['统一社会信用代码'],rawdata['统一社会信用代码_x'])
        rawdata['税收所属期'] = np.where(rawdata['税收所属期']!=rawdata['税收所属期'],rawdata['开票月份'],rawdata['税收所属期'])
        rawdata.fillna(0,inplace=True)
        
        
        custype=self.customer_re()
        dict1=dict(zip(custype['企业名称'],custype['官方类型']))
        dict2=dict(zip(custype['企业名称'],custype['类型']))
        result=pd.DataFrame()
        for i in range(rawdata.shape[0]):
            Series=rawdata.iloc[i]
            name=Series.企业名称
            if needregedit == False:
            
                if type(dict1.get(name))==str:
                    cs_type=dict1.get(name)
                else:
                    if type(dict2.get(name))==str:
                        cs_type=dict2.get(name)
                    else:
                        cs_type='能源类'
            else:
                if type(dict1.get(name))==str:
                    cs_type=dict1.get(name)
                else:
                    if type(dict2.get(name))==str:
                        cs_type=dict2.get(name)
                    else:
                        cs_type='未注册'
                
                    
            #_x总和，_y已申请，无标的为待申请
            Series['增值税.1_x']=Series['增值税_x']*self.df_7[self.df_7['类型']==cs_type]['增值税'].sum()
            Series['城建税.1_x']=Series['城建税_x']*self.df_7[self.df_7['类型']==cs_type]['城建税'].sum()
            Series['教育费附加.1_x']=Series['教育费附加_x']*self.df_7[self.df_7['类型']==cs_type]['教育费附加'].sum()
            Series['印花税.1_x']=Series['印花税_x']*self.df_7[self.df_7['类型']==cs_type]['印花税'].sum()
            Series['企业所得税.1_x']=Series['企业所得税_x']*self.df_7[self.df_7['类型']==cs_type]['企业所得税'].sum()
            Series['个人所得税.1_x']=Series['个人所得税_x']*self.df_7[self.df_7['类型']==cs_type]['个人所得税'].sum()
            Series['合计_x']=(Series['增值税.1_x']+Series['城建税.1_x']+Series['教育费附加.1_x']+Series['印花税.1_x']+Series['企业所得税.1_x']+Series['个人所得税.1_x']).sum()
            
            Series['增值税.1_y']=Series['增值税_y']*self.df_7[self.df_7['类型']==cs_type]['增值税'].sum()
            Series['城建税.1_y']=Series['城建税_y']*self.df_7[self.df_7['类型']==cs_type]['城建税'].sum()
            Series['教育费附加.1_y']=Series['教育费附加_y']*self.df_7[self.df_7['类型']==cs_type]['教育费附加'].sum()
            Series['印花税.1_y']=Series['印花税_y']*self.df_7[self.df_7['类型']==cs_type]['印花税'].sum()
            Series['企业所得税.1_y']=Series['企业所得税_y']*self.df_7[self.df_7['类型']==cs_type]['企业所得税'].sum()
            Series['个人所得税.1_y']=Series['个人所得税_y']*self.df_7[self.df_7['类型']==cs_type]['个人所得税'].sum()
            Series['合计_y']=(Series['增值税.1_y']+Series['城建税.1_y']+Series['教育费附加.1_y']+Series['印花税.1_y']+Series['企业所得税.1_y']+Series['个人所得税.1_y']).sum()
            
            Series['增值税.1']=Series['增值税']*self.df_7[self.df_7['类型']==cs_type]['增值税'].sum()
            Series['城建税.1']=Series['城建税']*self.df_7[self.df_7['类型']==cs_type]['城建税'].sum()
            Series['教育费附加.1']=Series['教育费附加']*self.df_7[self.df_7['类型']==cs_type]['教育费附加'].sum()
            Series['印花税.1']=Series['印花税']*self.df_7[self.df_7['类型']==cs_type]['印花税'].sum()
            Series['企业所得税.1']=Series['企业所得税']*self.df_7[self.df_7['类型']==cs_type]['企业所得税'].sum()
            Series['个人所得税.1']=Series['个人所得税']*self.df_7[self.df_7['类型']==cs_type]['个人所得税'].sum()
            Series['合计']=(Series['增值税.1']+Series['城建税.1']+Series['教育费附加.1']+Series['印花税.1']+Series['企业所得税.1']+Series['个人所得税.1']).sum()
            
            if cs_type == '能源类':
                Series['纳税合计']=Series['增值税_x']+Series['城建税_x']+Series['印花税_x']+Series['企业所得税_x']+Series['个人所得税_x']
                Series['营业收入1.5%']=Series['价税合计']*0.015
                Series['营业收入1.25%']=0
            if cs_type == '非能源类':
                Series['纳税合计']=Series['增值税_x']+Series['城建税_x']+Series['教育费附加_x']+Series['印花税_x']+Series['企业所得税_x']+Series['个人所得税_x']
                Series['营业收入1.5%']=0
                Series['营业收入1.25%']=Series['价税合计']*0.0125
            result=result.append(Series,ignore_index=True)   
        result['额度结余']=result['营业收入1.5%']+result['营业收入1.25%']-result['合计_x']
        result['累计结余']=np.nan
        result['是否保留']=np.nan
        
        usecopy=result.loc[:,['企业名称','税收所属期','额度结余']]    
        nameset=list(set(usecopy['企业名称']))
        for x in nameset:
            temp=usecopy[usecopy['企业名称']==x]
            temp.sort_values(by='税收所属期',axis=0,ascending=True,inplace=True)
            numlist=list(temp['额度结余'])
            numlist_raw=numlist
            savelist=['yes' for n in numlist]
            
                
            
            for y in range(1,len(numlist)):
                numlist[y]=numlist[y]+numlist[y-1]
            temp['累计结余']=numlist
            savelist=['yes' for n in numlist]
            
            if len(savelist)==1:
                if numlist[0]<0:
                    savelist[0]='no'
                else:
                    for e in range(1,len(savelist)):
                        if numlist[-e]>0:
                            print('a')
                            break
                        else:
                            print('b')
                            savelist[-e]='no'
            for q in range(len(savelist)):
                if savelist[q]=='no':
                    if numlist_raw[q]>=0:
                        savelist[q]='yes'
                        
            temp['是否保留']=savelist
                     
            
            
            result.update(temp['累计结余'])
            result.update(temp['是否保留'])
  
        result.drop(columns=['Unnamed: 0_x','Unnamed: 1_x','省级财力贡献_x','统一社会信用代码',
                             '市级财力贡献_x','统一社会信用代码_y','Unnamed: 0_y','Unnamed: 1_y',
                             '省级财力贡献_y','市级财力贡献_y','Unnamed: 0',
                             'Unnamed: 1','省级财力贡献','市级财力贡献','销方企业名称',
                             '客户编号','统一社会信用代码_z','开票月份','金额','税额','累计价税合计',
                             '销方企业所属','销方会员类型','份数','导入账户'],inplace=True)
        
  
        result_cut=result[result['合计']!=0]
        result_hand=result_cut[result_cut['是否保留']=='yes']
        if monthlimit==True:
            result_hand=result_hand[result_hand['额度结余']>=0]
        
        if confirm == True:
              mark_df=result_hand.loc[:,['统一社会信用代码_x','税收所属期','是否保留']]
              apply_con=pd.merge(needapply,mark_df,left_on=['Unnamed: 5','Unnamed: 15'],
                                 right_on=['统一社会信用代码_x','税收所属期'],how='left')
              ID_list=list(apply_con['Unnamed: 1'])
              apply_date=[self.today.strftime("%Y-%m-%d %H:%M:%S") for i in ID_list]
              new_apply=pd.DataFrame()
              new_apply['ID']=ID_list
              new_apply['apply_date']=apply_date
              self.df_applyed=self.df_applyed.append(new_apply,  ignore_index=True)
              self.df_applyed.to_excel('d:/DATABASE/all/applyed.xlsx')
              self.df_applyed.to_excel('d:/DATABASE/applyed/'+self.today.strftime("%Y-%m-%d")+'applyed.xlsx')

        if need == 'result':
            
            return result_hand 
        
        if need == 'raw':
            
            return result 
        if need == 'new':
            result_new=result.loc[:,['企业名称','统一社会信用代码_x','税收所属期','纳税合计','价税合计','合计_x','营业收入1.5%','营业收入1.25%','累计结余','额度结余','是否保留']]
            
            return result_new
    def customer_re(self):
        temp=self.Sale_calculate(start=202201,place='全部',method='type')
        tempsum=temp.groupby('统一社会信用代码').sum()
        typelist=[]
        for i in range(tempsum.shape[0]):
            if (tempsum.iloc[i,1]+tempsum.iloc[i,2])/tempsum.iloc[i].sum() > 0.4:
                typelist.append('非能源类')
            else:
                typelist.append('能源类')
        tempsum['类型']=typelist
        tempsum['认定时间']=self.tomonth

        tempsum.reset_index(inplace=True)
         
        cusdf=self.df_1.iloc[:,[1,4]]
        result=pd.merge(cusdf, tempsum,on='统一社会信用代码',how='left')
        result=self.add_name(result,1) 
        result.drop(columns=['energy','other','unenergy','unknow'],inplace=True)
        
        newDF=pd.merge(self.df_10, result,on='统一社会信用代码',how='outer')
        newDF.rename(columns={'企业名称_y':'企业名称'},inplace=True)
        newDF.to_excel('d:/DATABASE/Customertype/custype.xlsx')
        return newDF
    def feedback(self,a):
        if a==1:
            return self.df_1
        if a==2:
            return self.df_2
        if a==3:
            return self.df_3
        if a==4:
            return self.df_4
        if a==5:
            return self.df_5
        if a==6:
            return self.df_6
        if a==7:
            return self.namedict
        
    def add_type(self):

        df_2=self.df_2
        cargolist=list(self.df_2['商品和服务分类'])
        cargotype=[]
        for i in cargolist:

            for x in self.list_re_cargo:
                pattern = re.compile(r'(.*)'+x)
                result1 = pattern.search(str(i))
                if result1 != None:
                    cargotype.append(self.dict_type.get(x))
                    break
            if result1 == None:
                cargotype.append('unknow')

        df_2['type']=cargotype

        return df_2
    def add_name(self,pd,a):
        listname=[]
        if a==0:
            for i in pd.index:
                listname.append(self.namedict.get(i))
            pd.insert(0,'企业名称',listname)
        else:
            for i in list(pd['统一社会信用代码']):
                listname.append(self.namedict.get(i))
            pd.insert(0,'企业名称',listname)
        return pd
    def add_inform(self,raw_pd,on='统一社会信用代码',need=0):
       result=pd.merge(raw_pd, self.df_1,on=on,how='left')
       if need == 0:          
           return result
    def add_end_date(self,pd,a,on='统一社会信用代码'):
        listname=[]
        end_date_dict=dict(zip(self.df_1['统一社会信用代码'],self.df_1['停止合作时间']))
        if a==0:
            for i in pd.index:
                listname.append(end_date_dict.get(i))
            pd.insert(0,'停止合作时间',listname)
        else:
            for i in list(pd[on]):
                listname.append(end_date_dict.get(i))
            pd.insert(0,'停止合作时间',listname)
        return pd
    
    
    
    
    def add_tax_person(self,raw_pd,on='统一社会信用代码',left_on=''):
        tax_person=self.df_1.loc[:,['统一社会信用代码','企业名称','记账人/对接人','报税人/对接人','开票人/对接人']]
                                 
        result=pd.merge(raw_pd,tax_person,left_on=left_on,right_on=on,how='left')
        return result
        
        
    def add_sale(self,raw_pd,start,end,on,how='sum'):
        que_df=self.Sale_calculate(start=start,end=end)
        result=pd.merge(raw_pd,que_df,on=on,how='left')
        return result
        
    def add_tax(self,raw_pd,start,end,on,how='sum'):
        que_df=self.Sale_calculate(start=start,end=end)
        result=pd.merge(raw_pd,que_df,on=on,how='left')
        return result
    def add_tax_back(self,raw_pd,start,end,on,how='sum'):
        que_df=self.Sale_calculate(start=start,end=end)
        result=pd.merge(raw_pd,que_df,on=on,how='left')
        return result

        
        
    def Sale_calculate(self,start=202001,end=tomonth,place='全部',method='all',whole=True):
        if whole==False:
            df_3=self.df_3.drop(self.df_3[(self.df_3['销方企业名称']=='海南迈科供应链管理有限公司')|(self.df_3['销方企业名称']=='海南兴威供应链有限公司')|(self.df_3['销方企业名称']=='海南泰智有色金属有限公司')|(self.df_3['销方企业名称']=='智威（海南）新材料科技有限公司')].index)
            df_2=self.df_2.drop(self.df_2[(self.df_2['销方企业名称']=='海南迈科供应链管理有限公司')|(self.df_2['销方企业名称']=='海南兴威供应链有限公司')|(self.df_2['销方企业名称']=='海南泰智有色金属有限公司')|(self.df_2['销方企业名称']=='智威（海南）新材料科技有限公司')].index)   

        if whole==True:
            df_3=self.df_3
            df_2=self.df_2
       
        if place != '全部':
            df_3=df_3[df_3['注册地']==place]
            df_2=df_2[df_2['销方企业注册地']==place]
            
        if method=='all':
            mask=(df_3['开票月份'] >= start) & (df_3['开票月份'] <= end) 
            userange=df_3.loc[mask]
            result=userange.groupby(['统一社会信用代码']).sum()['价税合计']
            listname=[]
            for i in result.index:
                listname.append(self.namedict.get(i))
            dict_result={'统一社会信用代码' :result.index,'企业名称':listname,'价税合计':result.values}
            result1=pd.DataFrame(dict_result)
            return result1
    
        if method == 'goods':
            start_new=str(start)[:4]+'-'+str(start)[4:6]
            end_new=str(end)[:4]+'-'+str(end)[4:6]
            mask=(df_2['开票月份'] >= start_new) & (df_2['开票月份'] <= end_new) 
            userange=df_2.loc[mask]
            result=userange.groupby(['商品和服务分类','开票月份']).sum('价税合计')
            result=result.unstack()
            result.reset_index(inplace=True)
            result.drop(columns=['发票代码','发票号码','销方客户编号','含税单价（元）'],inplace=True)
            return result
        if method == 'type':
            start_new=str(start)[:4]+'-'+str(start)[4:6]
            end_new=str(end)[:4]+'-'+str(end)[4:6]
            mask=(df_2['开票月份'] >= start_new) & (df_2['开票月份'] <= end_new)
            userange=df_2.loc[mask]

            cargolist=list(userange['商品和服务分类'])
            cargotype=[]
            for i in cargolist:

                for x in self.list_re_cargo:
                    pattern = re.compile(r'(.*)'+x)
                    result1 = pattern.search(str(i))
                    if result1 != None:
                        cargotype.append(self.dict_type.get(x))
                        break
                if result1 == None:
                    cargotype.append('unknow')
            userange['type']=cargotype
            result=userange.groupby(['统一社会信用代码','开票月份','type']).sum()['价税合计（元）']
            result=result.unstack()
            result.reset_index(inplace=True)
            result=self.add_name(result, 1)
            
            return result

                
    def Tax_calculate(self,start=202001,end=tomonth,place='全部',method='all',whole=True): 
        if whole==False:
            try:
                df_4=self.df_4.drop(self.df_4[(self.df_4['公司名称']=='海南迈科供应链管理有限公司')|(self.df_4['公司名称']=='海南兴威供应链有限公司')|(self.df_4['公司名称']=='海南泰智有色金属有限公司')|(self.df_4['公司名称']=='智威（海南）新材料科技有限公司')].index)
            except:pass
        if whole==True:
            df_4=self.df_4
        
        if place=='海口':
            df_4=df_4[df_4['纳税地']==place] 
            mask=(df_4['所属月份'] >= start) & (df_4['所属月份'] <= end) 
            userange=df_4.loc[mask]
            if method=='all':            
                result=userange.groupby(['统一社会信用代码']).sum()['合计']
                listname=[]
                for i in result.index:
                    listname.append(self.namedict.get(i))
                dict_result={'统一社会信用代码' :result.index,'企业名称':listname,'纳税合计':result.values}
                result1=pd.DataFrame(dict_result)
                return result1 
        
            if method=='type':
                result=userange.groupby(['统一社会信用代码']).sum()
                
                result.drop(columns=['所属月份'],inplace=True)
                listname=[]
                for i in result.index:
                    listname.append(self.namedict.get(i))
                result.insert(0,'企业名称',listname)
                result=result.reset_index()
                return result
        if place=='洋浦':
            df_4=df_4[df_4['纳税地']==place] 
            mask=(df_4['所属月份'] >= start) & (df_4['所属月份'] <= end) 
            userange=df_4.loc[mask]
            if method=='all':            
                result=userange.groupby(['统一社会信用代码']).sum()['合计']
                listname=[]
                for i in result.index:
                    listname.append(self.namedict.get(i))
                dict_result={'统一社会信用代码' :result.index,'企业名称':listname,'纳税合计':result.values}
                result1=pd.DataFrame(dict_result)
                return result1 
            if method=='type':
               result=userange.groupby(['统一社会信用代码']).sum()
               
               result.drop(columns=['所属月份'],inplace=True)
               listname=[]
               for i in result.index:
                   listname.append(self.namedict.get(i))
               result.insert(0,'企业名称',listname)
               result=result.reset_index()
               return result
        else:
            mask=(df_4['所属月份'] >= start) & (df_4['所属月份'] <= end) 
            userange=df_4.loc[mask]
            if method=='all':            
                result=userange.groupby(['统一社会信用代码']).sum()['合计']
                listname=[]
                for i in result.index:
                    listname.append(self.namedict.get(i))
                dict_result={'统一社会信用代码' :result.index,'企业名称':listname,'纳税合计':result.values}
                result1=pd.DataFrame(dict_result)
                return result1 
            if method=='type':
                result=userange.groupby(['统一社会信用代码']).sum()
                
                result.drop(columns=['所属月份'],inplace=True)
                listname=[]
                for i in result.index:
                    listname.append(self.namedict.get(i))
                result.insert(0,'企业名称',listname)
                result=result.reset_index()
                return result
    def TaxBack_calculate(self,start=202001,end=tomonth,place='全部',method='all',whole=True): 
        if whole==False:
            try:
               df_41=self.df_41=self.df_41.drop(self.df_41[(self.df_41['会员企业名称']=='海南迈科供应链管理有限公司')|(self.df_41['会员企业名称']=='海南兴威供应链有限公司')|(self.df_41['会员企业名称']=='海南泰智有色金属有限公司')|(self.df_41['会员企业名称']=='智威（海南）新材料科技有限公司')].index)
            except:pass
        if whole==True:
            df_41=self.df_41           
        if place!='全部':
            df_41=df_41[df_41['纳税地']==place] 
        mask=(df_41['税收所属期'] >= start) & (df_41['税收所属期'] <= end) 
        userange=df_41.loc[mask]
        if method=='all':            
            result=userange.groupby(['统一社会信用代码']).sum()

            result.drop(columns=['税收所属期'],inplace=True)
            return result
        if method=='apply':
            result=userange.groupby(['统一社会信用代码']).sum()

            result.drop(columns=['税收所属期','纳税合计','增值税', '城建税','个税.2','应收','未收','Unnamed: 39',
                                 '教育费附加', '地方教育费附加', '印花税', '企业所得税', '个税', '实返合计', '增值税.2', '城建税.2',
                                 '教育费附加.2', '印花税.2', '企业所得税.2', '金额', '原因',  '已收'
                                  ],inplace=True)
            return result
        if method=='real':
            result=userange.groupby(['统一社会信用代码']).sum()

            result.drop(columns=['税收所属期','纳税合计', '增值税', '城建税','个税.1','Unnamed: 39',
                                 '教育费附加', '地方教育费附加', '印花税', '企业所得税', '个税', '申请合计', '增值税.1', '城建税.1',
                                 '教育费附加.1', '印花税.1', '企业所得税.1'],inplace=True)
            return result
                        

            
    def all_calculate(self,start=202001,end=tomonth,place='全部',method='all',whole=True):
        result_sale=self.Sale_calculate(start,end,place=place,method='all',whole=whole)
        result_tax=self.Tax_calculate(start,end,place=place,method='type',whole=whole)
        result_taxback_detail=self.TaxBack_calculate(start,end,place=place,method='all',whole=whole)
        result1=pd.merge(result_sale,result_tax,on='统一社会信用代码',how='outer')
        result=pd.merge(result1,result_taxback_detail,on='统一社会信用代码',how='outer')
        try:
            result.drop(columns=['纳税合计', '增值税', '城建税_y', '教育费附加_y', '地方教育费附加_y',
                                 '印花税_y', '企业所得税_y', '个税', '增值税.1', '城建税.1', '教育费附加.1', '印花税.1',
                                 '企业所得税.1', '增值税.2', '城建税.2', '教育费附加.2', '印花税.2', '企业所得税.2',
                                 '金额_y', '原因', '应付', '已收', 'Unnamed: 35' ],inplace=True)
        except:pass
        listname=[]
        bo=list(result['统一社会信用代码'])
        for i in bo:
            listname.append(self.namedict.get(i))
        result.insert(0,'企业名称',listname)
        result=result.reset_index()
        order=['统一社会信用代码','企业名称','价税合计','合计','实缴增值税','城建税_x','教育费附加_x','地方教育费附加_x','印花税_x','企业所得税_x','个人所得税','其他收入-工会经费','申请合计','实返合计']
        return result[order]
            
    def report(self,start=202001,end=tomonth,place='全部',method='all',whole=True):
        mask3=(self.df_3['开票月份'] >= start) & (self.df_3['开票月份'] <= end) 
        df_3=self.df_3.loc[mask3]
        mask4=(self.df_4['所属月份'] >= start) & (self.df_4['所属月份'] <= end) 
        df_4=self.df_4.loc[mask4]
        mask41=(self.df_41['税收所属期'] >= start) & (self.df_41['税收所属期'] <= end) 
        df_41=self.df_41.loc[mask41]
        if whole == False:
            try:
                df_3=df_3.drop(df_3[(df_3['销方企业名称']=='海南迈科供应链管理有限公司')|(df_3['销方企业名称']=='海南兴威供应链有限公司')|(df_3['销方企业名称']=='海南泰智有色金属有限公司')|(df_3['销方企业名称']=='智威（海南）新材料科技有限公司')].index)
                df_4=df_4.drop(df_4[(df_4['公司名称']=='海南迈科供应链管理有限公司')|(df_4['公司名称']=='海南兴威供应链有限公司')|(df_4['公司名称']=='海南泰智有色金属有限公司')|(df_4['公司名称']=='智威（海南）新材料科技有限公司')].index)
            except:pass
        if place== '海口':
            df_3=df_3[df_3['注册地']==place]
            df_4=df_4[df_4['纳税地']==place] 
            df_41=df_41[df_41['纳税地']==place] 
        if place== '洋浦':
            df_3=df_3[df_3['注册地']==place]
            df_4=df_4[df_4['纳税地']==place] 
            df_41=df_41[df_41['纳税地']==place] 
        result_sale=df_3.groupby(['开票月份']).sum().drop(columns=['客户编号','累计价税合计','份数','金额','税额'])
        result_tax=df_4.groupby(['所属月份']).sum()
        result_taxback_apply=df_41.groupby(['税收所属期']).sum().drop(columns=['纳税合计', '增值税', '城建税', '教育费附加', '地方教育费附加', '印花税', '企业所得税', '个税','城建税.2', '教育费附加.2', '印花税.2', '企业所得税.2', '金额', '原因',  '已收'])
        result_taxback_real=df_41.groupby(['税收所属期']).sum().drop(columns=['纳税合计', '增值税', '城建税', '教育费附加', '地方教育费附加', '印花税', '企业所得税', '个税','申请合计','增值税.1', '城建税.1', '教育费附加.1', '印花税.1', '企业所得税.1','金额', '原因',  '已收'])
        result_taxback_detail=df_41.groupby(['税收所属期']).sum()

        if method == 'all_detail':
            result=pd.merge(pd.merge(result_sale,result_tax,left_index=True,right_index=True,how='left'),result_taxback_real,left_index=True,right_index=True,how='left')
            return result
        if method == 'all':
            result=pd.merge(pd.merge(result_sale,result_tax,left_index=True,right_index=True,how='left'),result_taxback_detail,left_index=True,right_index=True,how='left')
            return result
        if method == 'sale':       
            return result_sale
        if method == 'tax':       
            return result_tax
        if method == 'taxback_apply':       
            return result_taxback_apply
        if method == 'taxback_real':       
            return result_taxback_real
        if method == 'taxback_detail':       
            return result_taxback_detail
  
    def rawmerge(self,df=pd.DataFrame(),start=202001,end=tomonth,place='全部',method='all',whole=True):
        df3=self.df_3
        df3['key']=df3['统一社会信用代码']+df3['开票月份'].astype('str')
        df4=self.df_4
        df4['key']=df4['统一社会信用代码']+df4['所属月份'].astype('str')
        df41=self.df_41
        df41['key']=df41['统一社会信用代码']+df41['税收所属期'].astype('str')
        result1=pd.merge(df3,df4,on='key',how='outer')
        result=pd.merge(result1,df41,on='key',how='outer')
        result['统一社会信用代码']=result['key'].str[:-6]
        bo=list(result['统一社会信用代码'])
        listname=[]
        for i in bo:
            listname.append(self.namedict.get(i))
        result.insert(0,'企业名称',listname)
        result['年月']=result['key'].str[-6:]
        
        return result
    
    def sale_rate(self,start='0000/00/00',end='0000/00/00',ID=''):
        rate_list=[]
        df_2=self.df_2
        start=start.replace('/','-')
        end=end.replace('/','-')
        mask=(df_2['开票日期'] >= start) & (df_2['开票日期'] <= end) & (df_2['统一社会信用代码'] == ID)
        userange=df_2.loc[mask]

        cargolist=list(userange['商品和服务分类'])

        cargotype=[]
        for i in cargolist:

            for x in self.list_re_cargo:
                pattern = re.compile(r'(.*)'+x)
                result1 = pattern.search(str(i))
                if result1 != None:
                    cargotype.append(self.dict_type.get(x))
                    break
            if result1 == None:
                cargotype.append('unknow')
        userange['type']=cargotype
        s1=userange.groupby(['type']).sum()['价税合计（元）']
        try:
            v0=sum(s1)
        except:v0=0
            
        try:
            v1=s1.energy
        except:v1=0
        try:
            v2=s1.unenergy
        except:v2=0
        try:
            v3=s1.other
        except:v3=0
        try:
            v4=s1.unknow
        except:v4=0
        default_list=[1,0,0,0,1]  #默认参数，第一位能源类占比，第二位非能源占比，第三位其他占比，第四位未知占比，第五位是否为缺省值
                                #缺省值为能源100%，缺省值为1
        if v0 == 0:
            return default_list
        else:
            rate_list=[v1/v0,v2/v0,v3/v0,v4/v0,0]
            return rate_list
                
  #洋浦税款计算默认按照税款所属期计算，可将datetype参数改为input，转为为按照入库期计算。（适用于按照开票明细拆分税款）
    def tax_cube(self,from_date=202201,end_date=tomonth,replace=True,need='',datetype='uninput'):
        rate_df=self.df_7
        YP_df=self.df_8   
       
        YP_df['Unnamed: 14'] = np.where(YP_df['Unnamed: 14']=='城市维护建设税','城建税', YP_df['Unnamed: 14'])
        YP_df['Unnamed: 14'] = np.where(YP_df['Unnamed: 14']=='车辆购置税','其他收入-工会经费', YP_df['Unnamed: 14'])
        YP_df['Unnamed: 14'] = np.where(YP_df['Unnamed: 14']=='城镇土地使用税','其他收入-工会经费', YP_df['Unnamed: 14'])
        YP_df['Unnamed: 14'] = np.where(YP_df['Unnamed: 14']=='房产税','其他收入-工会经费', YP_df['Unnamed: 14'])
        YP_df['Unnamed: 14'] = np.where(YP_df['Unnamed: 14']=='其他收入','其他收入-工会经费', YP_df['Unnamed: 14'])
        if datetype =='uninput':            
            mask=(YP_df['Unnamed: 15'] >= from_date) & (YP_df['Unnamed: 15'] <= end_date)
        if datetype =='input':
            mask=(YP_df['Unnamed: 13'] >= from_date) & (YP_df['Unnamed: 13'] <= end_date)
        userange=YP_df.loc[mask]
        result=pd.DataFrame()
        for i in range(len(userange)):
            Serie=userange.iloc[i]

        
            rate_s=self.sale_rate(start=Serie['起始日'],end=Serie['截止日'],ID=Serie['Unnamed: 5'])
            v1=Serie.iloc[16:23].sum()*rate_df[Serie.iloc[14]][1]*rate_s[0]
            v2=Serie.iloc[16:23].sum()*rate_df[Serie.iloc[14]][2]*rate_s[1]

            v0=v1+v2
            Serie['申请合计']=v0
            Serie['增值税.1']=0
            Serie['城建税.1']=0
            Serie['教育费附加.1']=0
            Serie['印花税.1']=0
            Serie['企业所得税.1']=0
            Serie['个人所得税.1']=0
            list_kind=['增值税','城建税','教育费附加','印花税','企业所得税','个人所得税']
            for i in list_kind:
                if i == Serie['Unnamed: 14']:
                    n=i+'.1'
                    Serie[n]=v0              
            Serie['能源类应返']=v1
            Serie['非能源类应返']=v2
            Serie['能源类占比']=rate_s[0]
            Serie['非能源类占比']=rate_s[1]
            Serie['不能返占比']=rate_s[2]
            Serie['未知占比']=rate_s[3]
            Serie['是否用了缺省值']=rate_s[4]
            result=result.append(Serie,ignore_index=True)   
            result.drop(columns=['Unnamed: 0'])
            
        if need == 'raw':        
            return result
        else:
            if datetype =='uninput': 
                result=result.groupby(['Unnamed: 5','Unnamed: 15']).sum()
                result.reset_index(inplace=True)
                result.rename(columns={'Unnamed: 5':'统一社会信用代码','Unnamed: 15':'税收所属期'},inplace=True)
                result=self.add_name(result,1)
            if datetype =='input':
                result=result.groupby(['Unnamed: 5','Unnamed: 13']).sum()
                result.reset_index(inplace=True)
                result.rename(columns={'Unnamed: 5':'统一社会信用代码','Unnamed: 13':'税收所属期'},inplace=True)
                result=self.add_name(result,1)
        if replace ==True:
            for i in range(len(result)):
                Serie=result.iloc[i]
                ID=Serie['统一社会信用代码']
                date=Serie['税收所属期']
                index_num=self.df_41[(self.df_41.统一社会信用代码==ID)&(self.df_41.税收所属期==date)].index.tolist()
                self.df_41.loc[index_num,'申请合计']= Serie['申请合计']
                self.df_41.loc[index_num,'增值税.1']= Serie['增值税.1']
                self.df_41.loc[index_num,'城建税.1']= Serie['城建税.1']
                self.df_41.loc[index_num,'教育费附加.1']= Serie['教育费附加.1']
                self.df_41.loc[index_num,'印花税.1']= Serie['印花税.1']
                self.df_41.loc[index_num,'企业所得税.1']= Serie['企业所得税.1']
                self.df_41.loc[index_num,'个税.1']= Serie['个人所得税.1']
             
        return result        
        
    #from :入库开始月份 end:入库截至月份 limit：税款所属期是否要在入库期所选范围内
    def tax_raw_cube(self,DF=pd.DataFrame(),from_date=202201,end_date=tomonth,limit=False,drop_close=False):
        if DF.shape[0]== 0:
            
            HK_df=self.df_81 
        
            star=str(from_date)[:4]+"/"+str(from_date)[4:6]+"/01"
            end=str(end_date)[:4]+"/"+str(end_date)[4:6]+"/31"
            mask=(HK_df['Unnamed: 13'] >= star) & (HK_df['Unnamed: 13'] <= end)
            userange=HK_df.loc[mask]
            if limit == True:
                mask2 = (HK_df['Unnamed: 15'] >= from_date) & (HK_df['Unnamed: 15'] <= end_date)
                userange=userange.loc[mask2]
        else:
            userange=DF
            
        if drop_close == True:
            userange=self.add_end_date(userange, 1,on='Unnamed: 5')
            userange['税款是否过服务期']=np.where(userange['Unnamed: 13']>userange['停止合作时间'],1,0)
            userange=userange[userange['税款是否过服务期']!= 1 ]
        result=userange.groupby(['Unnamed: 5','Unnamed: 15']).sum()
        result.reset_index(inplace=True)
        result.rename(columns={'Unnamed: 5':'统一社会信用代码','Unnamed: 15':'税收所属期'},inplace=True)
        result=self.add_name(result,1)
        return result        
    #计算收入 最老的一版，可以拆分开票明细
    def income(self,from_date=202201,end_date=tomonth,):
        self.tax_refresh()
        print('应返补正完成')
        mask=(self.df_41['税收所属期'] >= from_date) & (self.df_41['税收所属期'] <= end_date)
        userange=self.df_41.loc[mask]
        for i in range(len(userange)):
            Serie=userange.iloc[i]
            ID=Serie.统一社会信用代码
            date=Serie['税收所属期']
            apply_sum=Serie['申请合计']
            real_sum=Serie['实返合计']
            a=self.period_sale(date, ID)
            try:
                b=float(self.df_3[(self.df_3.统一社会信用代码==ID)&(self.df_3.开票月份==date)].价税合计)
            except:
                b=0
                print(ID+str(date)+'价税合计获取错误')
            if (a+b)<=500000000:
                n1=b
                n2=0
                n3=0
            else:
                if (a+b) <=1000000000:
                    n3=0
                    n2=a+b-500000000
                    if a <=500000000:
                        n1=500000000-a
                    else:
                        n1=0
                else:
                    n3=(a+b)-1000000000
                    if a >=1000000000:
                        n2=0
                        n3=0
                    else:
                        n2=1000000000-a
                        if a >=500000000:
                            n3=0
                        else:
                            n3=500000000-a
            if n1+n2+n3==0:
                rate=pd.Series([1,0,0])
            else:
                rate=pd.Series([n1/(n1+n2+n3),n2/(n1+n2+n3),n3/(n1+n2+n3)])
            rate2=pd.Series([0.1,0.05,0.03])
            income_sum=sum(apply_sum * rate * rate2)
            income_real_sum=sum(real_sum * rate * rate2)
            index_num=self.df_41[(self.df_41.统一社会信用代码==ID)&(self.df_41.税收所属期==date)].index.tolist()
            self.df_41.loc[index_num,'预计收入']= income_sum
            self.df_41.loc[index_num,'服务期内开票额']=a+b
            
            
            
        return self.df_41            
    
        
        

        for i in range(len(userange)):
            Serie=userange.iloc[i]
            ID=Serie.统一社会信用代码
            date=Serie['税收所属期']
            apply_sum=Serie['申请合计']
            real_sum=Serie['实返合计']
            a=self.period_sale(date, ID)
            try:
                b=float(self.df_3[(self.df_3.统一社会信用代码==ID)&(self.df_3.开票月份==date)].价税合计)
            except:
                b=0
                print(ID+str(date)+'价税合计获取错误')
            if (a+b)<=500000000:
                n1=b
                n2=0
                n3=0
            else:
                if (a+b) <=1000000000:
                    n3=0
                    n2=a+b-500000000
                    if a <=500000000:
                        n1=500000000-a
                    else:
                        n1=0
                else:
                    n3=(a+b)-1000000000
                    if a >=1000000000:
                        n2=0
                        n3=0
                    else:
                        n2=1000000000-a
                        if a >=500000000:
                            n3=0
                        else:
                            n3=500000000-a
            if n1+n2+n3==0:
                rate=pd.Series([1,0,0])
            else:
                rate=pd.Series([n1/(n1+n2+n3),n2/(n1+n2+n3),n3/(n1+n2+n3)])
            rate2=pd.Series([0.1,0.05,0.03])
            income_sum=sum(apply_sum * rate * rate2)
            income_real_sum=sum(real_sum * rate * rate2)
            index_num=self.df_41[(self.df_41.统一社会信用代码==ID)&(self.df_41.税收所属期==date)].index.tolist()
            self.df_41.loc[index_num,'预计收入']= income_sum
            self.df_41.loc[index_num,'服务期内开票额']=a+b
        
    
    #企业在当前日期，服务周期内的开票金额
    #date用整数类型，ID用统一社会信用代码
    def period_sale(self,date,ID):
        df_1=self.df_1
        df_3=self.df_3
        df_1.index=df_1['统一社会信用代码']
        temp=df_1.loc[ID,:]
        start_date_raw=temp.服务期限起始
        start_date_raw_1=temp['成立(迁移)日期']
        try:
            start_date=int(start_date_raw[:4]+start_date_raw[5:7])
        except:
            start_date=int(start_date_raw_1[:4]+start_date_raw_1[5:7])   
        zoom1=Interval(start_date,start_date+100,upper_closed=False)
        zoom2=Interval(start_date+100,start_date+200,upper_closed=False)
        zoom3=Interval(start_date+200,start_date+300,upper_closed=False)
        listA=[zoom1,zoom2,zoom3]
        for i in listA:
            if date in i :
                mask=(df_3['开票月份'] >= i.lower_bound) & (df_3['开票月份'] < date) &(df_3['统一社会信用代码'] == ID) 
                userange=df_3.loc[mask]
        try:
            result=sum(userange['价税合计'])
        except:
            result=0
            print (ID+'查询销售额错误')
        return result
    #按入库期计算税款
    def tax_calculate_input(self):
        pass
        
        
    
    #展示某个公司所有的信息
    def customer_all(self,name,ID,start,end):
        tax_result=self.all_calculate(start=start,end=end,place='全部',method='all',whole=True)
        
        pass
    

#查询匹配功能-用以索引每个公司当前的
    def raw_querry(self):
        print ('请将需要查询信息的公司在模板中输入>>>')
        df_que = pd.read_excel('d:/analyze program/查询匹配/查询列表.xlsx')
        tomonth=int(datetime.datetime.now().strftime('%Y%m'))
        DATA=initialize()
        df_1=DATA.feedback(3)
        DF_trade=df_1[df_1['开票月份']==tomonth]
        sum_dict=dict(zip(DF_trade['销方企业名称'],DF_trade['累计价税合计']))
        namelist=df_que['公司名称']
        resultlist=[]
        count1,count2= 0,0
        for i in namelist:
            try:
                resultlist.append(sum_dict[i])
                count1+=1
            except:
                resultlist.append('未查询到')
                count2+=1
        result=pd.DataFrame({'公司名称':namelist,'累计价税合计':resultlist})
        
        df_2=DATA.feedback(4)
        tax_result=df_2.groupby(['公司名称']).sum(['合计','实缴增值税','城建税','教育费附加','地方教育费附加','印花税','企业所得税','个人所得税','其他收入-工会经费'])
        result=pd.merge(result, tax_result.drop(columns=['所属月份']),how='left',on='公司名称')
        
        count=count1+count2
        print(result)
        print('querry finished !')
        print('total '+str(count))
        print('Find '+str(count1))
        print('No match '+str(count2))
        result.to_excel('d:/analyze program/匹配结果/匹配结果.xlsx',index=None)
       
    def querry(start,end):
        print ('请将需要查询信息的公司在模板中输入>>>')
        print ('查询的时间区间为'+str(start)+'>>>'+str(end))
        df_que = pd.read_excel('d:/analyze program/查询匹配/查询列表.xlsx')
        DATA=initialize()
        df_1=DATA.feedback(3)
        DF_trade=df_1[df_1['开票月份']==end]
        DF_trade_star=df_1[df_1['开票月份']==start]
        sum_dict=dict(zip(DF_trade['销方企业名称'],DF_trade['累计价税合计']-DF_trade_star['累计价税合计']+DF_trade_star['价税合计']))
        namelist=df_que['公司名称']
        resultlist=[]
        count1,count2= 0,0
        for i in namelist:
            try:
                resultlist.append(sum_dict[i])
                count1+=1
            except:
                resultlist.append('未查询到')
                count2+=1
        result=pd.DataFrame({'公司名称':namelist,'累计价税合计':resultlist})
        
        df_raw=DATA.feedback(4)
        mask=(df_raw['所属月份'] >= start) & (df_raw['所属月份'] <= end)
        df_2=df_raw.loc[mask]
        tax_result=df_2.groupby(['公司名称']).sum(['合计','实缴增值税','城建税','教育费附加','地方教育费附加','印花税','企业所得税','个人所得税','其他收入-工会经费'])
        result=pd.merge(result, tax_result.drop(columns=['所属月份']),how='left',on='公司名称')
        
        count=count1+count2
        print(result)
        print('querry finished !')
        print('total '+str(count))
        print('Find '+str(count1))
        print('No match '+str(count2))
        result.to_excel('d:/analyze program/匹配结果/匹配结果.xlsx',index=None)
    

    
#日常查询数据

    


#while(True):
#    print('请输入功能:0.自动下载数据 1.数据报表 2.报告制作 3.数据分析 9.退出')
#    a=input()
#    if a=='0':
#        at.autoDownload()


