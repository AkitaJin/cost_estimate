#coding:utf-8

import os
import numpy as np
import pandas as pd
import psycopg2
from sqlalchemy import create_engine


"""
本代码通过技术科提供基准介对应的图号,并将该图号BOM录入SageX3-ERP中，
对应ERP的采购入库材料价格,获得该BOM的材料及其物料价格。
1。 输入
1.1 BOM清单
1.2 物料替代清单
1.3 最新物料入库价格表
1.4 读取所有图号的BOM集合
"""

# 一. 读取所有数据输入源

# 1. 读入最新物料价格
material = pd.read_excel(r"..\src\核价工具11.0\最新入库价(7).xlsx",
                            sheet_name="Sheet1",
                            engine="openpyxl")

# 2. 读入对照表,作用是区别主材与其他材料
material_matched = pd.read_excel(r"G:\数据中台\05-核价及成本\核价工具11.0\对照表.xlsx",
                   sheet_name="Sheet2",
                   engine="openpyxl")
material_matched.drop_duplicates(subset=["替代物料识别辅助列"], keep='first', inplace=True)

# 3 读入技术科提供的图号清单
BOM_drawing = pd.read_excel(r"G:\数据中台\05-核价及成本\核价工具11.0\所有图号清单.xlsx",
                            sheet_name="Sheet1",
                            engine="openpyxl")

# 4 读取物料替代清单<材料替代清单.xlsx>
BOM_Replace = pd.read_excel(r"G:\数据中台\05-核价及成本\核价工具11.0\材料替代清单.xlsx",
                            sheet_name="Sheet1",
                            engine="openpyxl")

# 1.3 读取PG库中所有图号的BOM清单并改列名
BOM_All = pd.read_excel(r"G:\数据中台\05-核价及成本\核价工具11.0\所有图号BOM清单.xlsx",
                            sheet_name="Sheet1",
                            engine="openpyxl")

BOM_All = BOM_All.iloc[:,1:]
col_1 = BOM_All[1:].columns.tolist()
print(col_1)
col_2 = ["图号编码","地点","图号名称","图号","材料编码","材料名称","材料代号",
         "库存数量","库存单位","创建日期","单位","采购数量","采购单位"]
# 组成字典,用于改列名
col_name =  dict(zip(col_1, col_2))
print(col_name)
BOM_All.rename(columns= col_name,inplace=True)
BOM_All["校准前"] = BOM_All["材料名称"] + BOM_All["材料代号"]

# 分组
groupby_df = BOM_All.groupby(['图号'])
u=[]
for i in groupby_df:
    # 替换材料代码
    i[1].to_excel(r"C:\Users\Administrator\Desktop\BOM-1\{0}.xlsx".format(i[0]))
    # 匹配物料替代表,为下一步匹配ERP最新物料价格做准备
    df_macth = BOM_Replace[["校准前","校准后","校准后物料编码","策略"]]
    df = i[1].merge(df_macth,how="left",on=["校准前"],validate="many_to_one")
    df["校准后物料编码"] = df["校准后物料编码"].fillna(df["材料编码"])
    df["校准后"] = df["校准后"].fillna(df["校准前"])
    # 匹配<对照表>,获取主材的大,中,小类及编码,为匹配主材价格做好准备工作
    df = df.merge(material_matched,how="left",left_on=["校准后"],right_on=["替代物料识别辅助列"])
    # 匹配<最新价格表>
    material = material[["材料编码","库存单位","库存单价"]]
    df = df.merge(material,how="left",left_on=["校准后物料编码"],right_on=["材料编码"])


    u.append(df)
    df.to_excel(r"C:\Users\Administrator\Desktop\BOM-2\{0}.xlsx".format(i[0]))
df4 = pd.concat(u)

# 所有合并后的图号材料清单
df4.to_excel(r"C:\Users\Administrator\Desktop\{0}.xlsx".format("AAA"))
