import pandas as pd

df = pd.read_excel('G:/Other computers/My Computer/Data [F]/MOD/2021\May-2021.xlsx', sheet_name='QF-LDC-33', skiprows=6,
                                   nrows=160, usecols='B:F', index_col=None, header=None)
print(df)
