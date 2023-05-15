# -*- coding: utf-8 -*-
"""
Created on Tue Sep  6 14:05:08 2022

@author: LV
"""

from sqlalchemy import create_engine
import pymysql
engine=create_engine("mysql+pymysql://root:893104@localhost:3306/data")
import pandas as pd
import newCUSRP as CUSRP

data=CUSRP.initialize()

#sql_db=pd.read_sql('test',con=engine)
data.df_1.to_sql(name='memberbook',con=engine,)
data.df_2.to_sql(name='Sale_detail',con=engine)
data.df_3.to_sql(name='Sale_report',con=engine)
data.df_4.to_sql(name='tax',con=engine)
data.df_41.to_sql(name='tax_back',con=engine)
data.df_7.to_sql(name='tax_rate',con=engine)
data.df_8.to_sql(name='tax_detail',con=engine)
data.df_9.to_sql(name='cargo_type',con=engine)






