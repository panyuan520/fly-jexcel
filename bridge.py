#!/usr/bin/env python
# -*- coding: utf-8 -*-
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