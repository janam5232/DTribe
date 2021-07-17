import numpy as np
import sqlite3 as sql3
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
    # if(len(fileNamesList[i]) == 3 or len(fileNamesList[i]) == 2):
    #     manyEMRMDF.append(pd.read_excel(f, sheet_name = None))
    #     print(fileNamesList[i])
    i += 1

# for f in files:
#     databaseDF.append(pd.read_excel(f, sheet_name = ('Data')))
#     print(databaseDF[i])
#     i = i + 1
# while(i<len(files)):

#SECTION RELATED TO STORING DATA SHEET TO SQLITE DATABASE

columnsToBeExtractedFromDataSheet = ['Date', 'FacilityType', 'BedSize', 'Region', 'Manufacturer', 'Ticker', 'Group', 'Therapy', 'Anatomy','SubAnatomy', 'ProductCategory', 'Quantity', 'AvgPrice', 'TotalSpend']
dbName = 'Data'
conn = sql3.connect(dbName + '.sqlite')
cur = conn.cursor()

for sheets in oneEMRMDF:
    print(sheets['Data'].isna().sum())
    df = sheets['Data'][columnsToBeExtractedFromDataSheet]

    df.to_sql(name='TableData', con=conn, if_exists='append')

# cur.execute('SELECT * FROM TableData')
# names = list(map(lambda x: x[0], cur.description))
# print(names)
# for row in cur:
#     print(row)
# cur.close()
