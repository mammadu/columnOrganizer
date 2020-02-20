# -*- coding: utf-8 -*-
"""
Created on Thu Feb 13 22:13:18 2020

@author: muham
"""

from excelOrganizer import excelOrganizer
from openpyxl import load_workbook

# fileLocation = input("type in the file location")

#locations of the files
# file1 = 'W:\\Documents\\Code\\excelOrganizer\\Block data.xlsx'
# file2 = 'W:\\Documents\\Code\\excelOrganizer\\Phase 3 usage.xlsx'

file1 = 'Block data.xlsx'
file2 = 'Phase 3 usage.xlsx'

#instantiating a class
eo = excelOrganizer()

#converting columns from excel files into dataframes.
c1 = eo.copyColumn(file1,'Current Carbon block part number') #this is the key column
c2 = eo.copyColumn(file2,'Current Carbon block part number') # this is the column to be evaluated
n1 = eo.copyColumn(file1,'new carbon block part number') #this column contains the data to be copied
n2 = eo.copyColumn(file2,'new carbon block part number') #this is the column that will be copied over

# print(n2)

#if the value in (file1,carbonBlockColumn) equals the value in (file2,carbonBlockColumn) replace the value of (file2,newColumn) with the value of (file1,newColumn)
for i in range(0,len(c1)): #this loop is to go through every cell in in the first file
    for j in range(0,len(c2)): #this loope evaluates the value in the first file and then makes a change to the new column in the second file.
        if c1[i]==c2[j]:
            n2[j] = n1[i]


wb = load_workbook('Phase 3 usage.xlsx') #this loads the workbook
ws = wb['Phase III with usage'] #this specifies the worksheet in the workbook

for index, row in n2.iterrows():
    cell = 'F%d'  % (index + 2) #this specifies the specific cell to edit. %d is a placeholder for a number which is given by %()


n2.to_excel(file2, sheet_name='Phase III with usage',index_label = 'new carbon block part number')
#copy n2 dataframe into phase 3 usage excel sheet.