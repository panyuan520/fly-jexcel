#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os
from bridge import *
import threading
import datetime

filename     = os.path.join(os.path.abspath(os.path.curdir), "result.xls")
templatePath = os.path.join(os.path.abspath(os.path.curdir), "template.xls")

header = [
         ["A1", [["SO#","PO#","Supplier Short Name", "designation","Set Item","PO Qty",'remark']]], #生成表格头
         ["A2", [
                    [1,2,3,4,5,6,7],
                    [1,2,3,4,5,6,7],
                    [1,2,3,4,5,6,7],
                    [1,2,3,4,5,6,7],
                    [1,2,3,4,5,6,7],
                 ]], #生成表格数据
          ["B40", "test path"],
          ["A10", "test align"]
         ]
         
'''
#批量测试数据
for i in range(10000):
    header.append(
                [20 + i, 0, [
                    [1,2,3,4,5,6,7],
                 ]], #生成表格数据
                 )
'''
#描述类似于CSS样式，可在java setStyle里面扩展                   
style = {
        "A11:G16":{"border":"solid", "borderColor":"black"}, #根据表格描线
        "B7":{'excute':"SUM(B3:B5)"},#求和
        "B10:D10":{'merged':'True'},
        "B10:D10":{'align':'top'},
        "G13":{'align':'center'},
         }
        
bridge = Bridge()
header = bridge.format(header)#格式化数据 将python数据类型转成java数据类型，可扩展
style = bridge.format(style)
bridge.report(templatePath, filename , header, style)
