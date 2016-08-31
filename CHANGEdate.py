# -*- coding: utf-8 -*-
"""
Created on Fri Jul 29 20:32:06 2016

@author: Administrator
"""
reload(sys)
sys.setdefaultencoding("utf-8") # python默认是以ascii进行编解码，跟"coding: UTF-8 "头无关
import xlrd  
import re  
import sqlite3  
import xlwt
import pandas as pd
import numpy as np
df = pd.read_excel('G:\CAICT\hang.xlsx')

workbook = xlrd.open_workbook('G:\CAICT\hang.xlsx')
booksheet = workbook.sheet_by_name('s1')
i=1
DATE1=range(booksheet.ncols)
DATE2=range(booksheet.ncols)
for cols in range(booksheet.ncols):
    cel=booksheet.cell(4, cols)
    val=cel.value  
    patternF = re.compile(r'为 *')
    patternB = re.compile(r' *人 *')
    matchF = patternF.search(val) 
    matchB = patternB.search(val) 
   
    if matchF :
        i+=i
        print matchF.end()
        indexFday1= matchF.end()              
        print matchB.end()
        indexBday1= matchB.start()
           
        day1=val[indexFday1:indexBday1]#获得起始日期
        indexday=val.find('/')        
        DAY=val[indexFday1:indexday]        
        indexmonth=val.find('/',(indexday+1))
        MON=(val[(indexday+1):indexmonth])
        print (MON)#获得起始日期的月                
        YEAR=(val[(indexmonth+1):indexBday1])
        print (YEAR)#获得起始日期的年        
        DATE1[row] = YEAR+'/'+MON+'/'+DAY+'/'
        print (DATE1[row])
        
        indexFday2= matchB.end()
        indexBday2=val.find('/',indexFday2) 
        DAY2=val[indexFday2:indexBday2]        
        indexmonth=val.find('/',(indexBday2+1))
        MON2=(val[(indexBday2+1):indexmonth])        
        YEAR2=(val[(indexmonth+1):])
        DATE2[row] = YEAR2+'/'+MON2+'/'+DAY2+'/'
        print (DATE2[row])  
        
#        booksheet.write(row, 9,DATE1[row])
#        booksheet.write(row, 9,DATE2[row])
        
    else:
        DATE1[row] = 'NaN'
        DATE2[row] = 'NaN'
        
    
        
        
        
        
        
        
        