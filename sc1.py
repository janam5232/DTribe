import pandas as pd
import os
import glob as gb

path = os.getcwd()
files = gb.glob(os.path.join(path, "*.xlsx"))

for f in files:
    df = pd.read_excel(f, sheet_name="Regression Model")
    print("Location: ", f)
    print('File Name: ', f.split("\\")[-1])

    print('Content: ')
    print(df.head())