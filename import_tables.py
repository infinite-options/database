#author: Purva Deekshit
#Created: August 2, 2019
#Modified: August 6, 2019
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile


#print(pd.__version__)
df_1 = pd.read_excel('Database Project.xlsx', sheet_name='INVENTORY')

print("Column headings:")
print(df_1.columns)
print()
print(df_1.head(10))

df_2 = pd.read_excel('Database Project.xlsx', sheet_name='BOM')

print("\nColumn headings:")
print(df_2.columns)
print()
print(df_2.head())