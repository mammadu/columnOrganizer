# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
fileLocation = r'C:\Users\E1231495\Desktop\coding\Phase 3 usage.xlsx'
df = pd.read_excel(fileLocation)
z = df.loc[:,'Current Carbon block part number']
print(z)

#I need to open both files, and somehow transplant the new carbon block part number column to my phase 3 usage excel file.