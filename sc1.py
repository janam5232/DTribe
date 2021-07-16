import numpy as np
import pandas as pd
import glob as gb
import os

#getting the current directory's location
path = os.getcwd()

#getting the name of all the excel files
files = gb.glob(os.path.join(path, "*.xlsx"))

# name = r"^R{1}[a-zA-Z\s]+$"
dictionaryOfdataframesOfExcelFiles = {}
fileNamesList = []
oneEMRMDF = []
manyEMRMDF = []
databaseDF = []
i = 0

#looping through the directory and extracting all the excel files' data
for f in files:
    file_name = f.split('\\')[-1]
    fileNamesList.append(file_name.split()[0])
    print(f)
    if(len(fileNamesList[i]) == 4):
        oneEMRMDF.append(pd.read_excel(f, sheet_name = (["Regression Model", "Empirical Model", "Data"])))
        print(fileNamesList[i])
    if(len(fileNamesList[i]) == 3 or len(fileNamesList[i]) == 2):
        manyEMRMDF.append(pd.read_excel(f, sheet_name = None))
        print(fileNamesList[i])
    i += 1

# for f in files:
#     databaseDF.append(pd.read_excel(f, sheet_name = ('Data')))
#     print(databaseDF[i])
#     i = i + 1
# while(i<len(files)):
