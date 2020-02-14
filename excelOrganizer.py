# -*- coding: utf-8 -*-
"""
Spyder Editor

This program is to copy data between excel files based on column and row location and information

Functions:
    copyPaste(self, file1, file 2)
        copy column data from between excel files
    
    

"""



import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile

fileLocation = r'W:\Documents\Code\excelOrganizer\excelOrganizer\Block data.xlsx'

class excelOrganizer(object):
    
    def copyColumn(fileLocation,columnLabel): #todo: possibly add in way to select specific rows
        df = pd.read_excel(fileLocation)
        z = df.loc[:,columnLabel]
        return z



#I need to open both files, and somehow transplant the new carbon block part number column to my phase 3 usage excel file.