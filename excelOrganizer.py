# -*- coding: utf-8 -*-
"""
Spyder Editor

This program is to copy data between excel files based on column and row location and information

Functions:
    copyPaste(self, file1, file 2)
        copy column data from between excel files
    
    

"""



import pandas as pd

#some guide I was reading said to insert the below code, but I have yet to use it. I may get rid of it.
# from pandas import ExcelWriter
# from pandas import ExcelFile


class excelOrganizer(object):
    
    #this below code is if I want to make the file location an attribute of the class.
    # def __init__(self, file1, file2):
    #     self.file1 = file1
    #     self.file2 = file2
    
    #this method reads and excel file at a location and converts a column into a dataframe.
    def copyColumn(self, fileLocation, columnLabel): #todo: possibly add in way to select specific rows
        df = pd.read_excel(fileLocation)
        z = df.loc[:,columnLabel]
        return z
    
    # def pasteColumn(self,fileLocation,columnLabel, data):
    #     pd.DataFrame.to_excel(fileLocation, columns=data, index_label = columnLabel)



#I need to open both files, and somehow transplant the new carbon block part number column to my phase 3 usage excel file.