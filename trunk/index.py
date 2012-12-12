# coding: utf-8
import os
import decimal
from datetime import datetime as dt
from py4j.java_gateway import JavaGateway

class Bridge(object):
    
    def __init__(self):
        gateway = JavaGateway()
        self.point = gateway.entry_point
        
    def format(self, obj):
        
        if obj !=0 and not obj:
            return ''
            
        if isinstance(obj, (int, float, long, basestring)):
            return obj
        
        elif isinstance(obj, dt):
            return str(obj)
        
        elif isinstance(obj, decimal.Decimal):
            return float(obj)
        
        elif isinstance(obj, list):
            jlist = self.point.create_list()
            for obj1 in obj:
                jlist.add(self.format(obj1))
            return jlist
        
        elif isinstance(obj, dict):
            jhashmap = self.point.create_hashmap()
            for key, obj1 in obj.iteritems():
                jhashmap.put(key, self.format(obj1));
            return jhashmap
        
    def report(self, im, ou, header, style):
        return self.point.report(im, ou, header, style)

        
filename     = os.path.join(os.path.abspath(os.path.curdir), "result.xls")
templatePath = os.path.join(os.path.abspath(os.path.curdir), "template.xls")

header = [
         [2, 1, 1232], #������,����ᣬ���� 
         [3, 1, 3424], 
         [4, 1, 1.22112], 
         [6, 0, 'Total'], 
         [10, 0, [["SO#","PO#","Supplier Short Name", "designation","Set Item","PO Qty",'remark']]], #���ɱ��ͷ
         [11, 0, [
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
         
import threading
import datetime

bridge = Bridge()
header = bridge.format(header)#��ʽ������ ��python��������ת��java�������ͣ�����չ
style = bridge.format(style)
bridge.report(templatePath, filename , header, style)
'''
class ThreadClass(threading.Thread):
  def run(self):
    bridge.report(templatePath, "%s.xls" % self.getName(), header, style)
    now = datetime.datetime.now()
    print "%s says Hello World at time: %s" % (self.getName(), now)
print dt.now()
for i in range(0):
  t = ThreadClass()
  t.start
'''          
