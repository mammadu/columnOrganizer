# -*- coding: utf-8 -*-
"""
Created on Thu Feb 13 22:13:18 2020

@author: muham
"""

from excelOrganizer import excelOrganizer as eo

x = eo.copyColumn(r'W:\Documents\Code\excelOrganizer\excelOrganizer\Block data.xlsx','Current Carbon block part number')
y = eo.copyColumn(r'W:\Documents\Code\excelOrganizer\excelOrganizer\Phase 3 usage.xlsx','Current Carbon block part number')
z = eo.copyColumn(r'W:\Documents\Code\excelOrganizer\excelOrganizer\Block data.xlsx','new carbon block part number')
w = eo.copyColumn(r'W:\Documents\Code\excelOrganizer\excelOrganizer\Phase 3 usage.xlsx','new carbon block part number')

print(w)

for i in range(0,len(x)):
    for j in range(0,len(y)):
        if x[i]==y[j]:
            w[i] = z[j]
        
print(w)


#determine why you are gettign 'KeyError: 62'
#copy w dataframe into phase 3 usage excel sheet.