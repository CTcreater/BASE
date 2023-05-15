# -*- coding: utf-8 -*-
"""
Created on Tue Sep  6 10:10:54 2022

@author: LV
"""

import pandas as pd
from sqlalchemy import create_engine

# 创建了一个mysql的工具类，方便使用
# df_write_mysql -> DataFrame write into mysql xx database xx table function
class MySQLUtil:
    # host：ip地址，port：端口号，username：用户名，password：密码，db：数据库名，table：表名
    def __init__(self, host='127.0.0.1', port='3306', username='root', password='893104', db='new_schema', table):
        self.host = host
        self.port = port
        self.username = username
        self.password = password
        self.db = db
        self.table = table
        # write_mode：操作方式   方式有：append,fail,replace
        #append：如果表存在，则将数据添加到表后面
        #fail：如果表存在就不操作
        #replace：如果表存在，删了，覆盖重建
        self.write_mode = 'replace'
        # 链接键格式 mysql+pymysql://用户名:密码@数据库地址/数据库名?charset=utf8’
        self.connect_url = 'mysql+pymysql://' + username + ':' + password + '@' + host + ':' + port + '/' + db + '?charset=utf8'
        self.mysql_connect = create_engine(self.connect_url)

    def df_write_mysql(self, data):
        # 参数设置:DataFrame 表名 链接键 数据库名 操作方式 是否录入索引
        pd.io.sql.to_sql(data, self.table, self.mysql_connect, schema=self.db, if_exists=self.write_mode, index=False)
        print("write into mysql finish")

