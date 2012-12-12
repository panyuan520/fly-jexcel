#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os
from bridge import *
import threading
import datetime

filename     = os.path.join(os.path.abspath(os.path.curdir), "result.xls")
templatePath = os.path.join(os.path.abspath(os.path.curdir), "template.xls")

header = [
         ["A1", [["SO#","PO#","Supplier Short Name", "designation","Set Item","PO Qty",'remark']]], #���ɱ��ͷ
         ["A2", [
                    [1,2,3,4,5,6,7],
                    [1,2,3,4,5,6,7],
                    [1,2,3,4,5,6,7],
                    [1,2,3,4,5,6,7],
                    [1,2,3,4,5,6,7],
                 ]], #���ɱ������
          ["B40", "test path"],
          ["A10", "test align"]
         ]
         
'''
#������������
for i in range(10000):
    header.append(
                [20 + i, 0, [
                    [1,2,3,4,5,6,7],
                 ]], #���ɱ������
                 )
'''
#����������CSS��ʽ������java setStyle������չ                   
style = {
        "A11:G16":{"border":"solid", "borderColor":"black"}, #���ݱ������
        "B7":{'excute':"SUM(B3:B5)"},#���
        "B10:D10":{'merged':'True'},
        "B10:D10":{'align':'top'},
        "G13":{'align':'center'},
         }
        
bridge = Bridge()
header = bridge.format(header)#��ʽ������ ��python��������ת��java�������ͣ�����չ
style = bridge.format(style)
bridge.report(templatePath, filename , header, style)
