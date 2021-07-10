#imported libraries
import pandas as pd
import os
import glob as gb

#getting the current directory's location
path = os.getcwd()
#getting the name of all the excel files
files = gb.glob(os.path.join(path, "*.xlsx"))

#looping through the directory and extracting all the excel files' data
for f in files:
    df = pd.read_excel(f, sheet_name="Regression Model")
    print("Location: ", f)
    print('File Name: ', f.split("\\")[-1])

    print('Content: ')
    print(df.head())
