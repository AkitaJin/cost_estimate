# -*- coding: utf-8 -*-
# ---
# jupyter:
#   jupytext:
#     formats: ipynb,py:light
#     text_representation:
#       extension: .py
#       format_name: light
#       format_version: '1.5'
#       jupytext_version: 1.11.4
#   kernelspec:
#     display_name: Python 3
#     language: python
#     name: python3
# ---

# +
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
path = r"J:\python_code\cost_estimate\src\核价工具11.0\\"
file_material = r"最新入库价(7).xlsx"
file_material_matched = r"对照表.xlsx"
file_BOM_drawing = r"所有图号清单.xlsx"
file_BOM_Replace = r"材料替代清单.xlsx"
file_BOM_All = r"所有图号BOM清单.xlsx"

# 1. 读入最新物料价格

material = pd.read_excel(path+file_material,
                            sheet_name="Sheet1"
                            # ,engine="openpyxl"
                            )
# print(material)

# 2. 读入对照表,作用是区别主材与其他材料
material_matched = pd.read_excel(path+file_material_matched,
                   sheet_name="Sheet2"
                #    ,engine="openpyxl"
                   )
material_matched.drop_duplicates(subset=["替代物料识别辅助列"], keep='first', inplace=True)

# 3 读入技术科提供的图号清单
BOM_drawing = pd.read_excel(path+file_BOM_drawing,
                            sheet_name="Sheet1"
                            # ,engine="openpyxl"
                            )

# 4 读取物料替代清单<材料替代清单.xlsx>
BOM_Replace = pd.read_excel(path+file_BOM_Replace,
                            sheet_name="Sheet1"
                            # ,engine="openpyxl"
                            )

# 1.3 读取PG库中所有图号的BOM清单并改列名
BOM_All = pd.read_excel(path+file_BOM_All,
                            sheet_name="Sheet1"
                            # ,engine="openpyxl"
                            )
# -

BOM_All = BOM_All.iloc[:,1:]

col_1 = BOM_All[1:].columns.tolist()

print(col_1)

col_2 = ["地点","图号编码","图号名称","图号","材料编码","材料名称","材料代号","库存数量","库存单位","创建日期","单位","采购数量","采购单位"]
# 组成字典,用于改列名
col_name =  dict(zip(col_1, col_2))

print(col_name)

BOM_All.rename(columns= col_name,inplace=True)

# +
tmp = BOM_All.copy()

tmp["校准前"] = tmp["材料名称"].astype(str) + tmp["材料代号"].astype(str)
# set(tmp["材料名称"])
# -








