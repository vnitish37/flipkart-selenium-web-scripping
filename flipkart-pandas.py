import pandas as pd 
NAME = input("Name of Excel file :\n")
data =pd.read_excel("{}.xlsx".format(NAME))
print(data)

