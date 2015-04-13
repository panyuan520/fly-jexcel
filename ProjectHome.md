
```
'''
Example:
开发了系统有很多导出excel报表，最开始用pywin32com，但是发现其中有各种各样的问题，
转向用java来解决问题，为了偷懒，写了一个方法不需要写java，也不需要写各种excel生成方法
来生成excel。
simple python bridge java poi export excel
python 调用java poi 生成 excel
根据数据定位和定位样式来生成需要的excel
'''

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
for i in range(10000):
    header.append(
                [20 + i, 0, [
                    [1,2,3,4,5,6,7],
                 ]], #生成表格数据
                 )
'''
#描述类似于CSS样式，可在java setStyle里面扩展                   
style = {
        "A1:G16":{"border":"solid", "borderColor":"black"}, #根据表格描线
        "B7":{'excute':"SUM(B3:B5)"},#求和
        "B10:D10":{'merged':'True'},#合并单元格
        "B10:D10":{'align':'top'},#设置表格中的位置
        "G13":{'align':'center'},
        "sheet":{0:"sheet0"}
         }
        
bridge = Bridge()
header = bridge.format(header)#格式化数据 将python数据类型转成java数据类型，可扩展
style = bridge.format(style)
bridge.report(templatePath, filename , header, style)





```

