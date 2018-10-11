import os
import pandas as pd
path = os.getcwd()
files = os.listdir(path)
files_xlsx = [f for f in files if f[-4:] == 'xlsx']
df = pd.DataFrame()
for f in files_xlsx:
    data = pd.read_excel(f, 'Validation')
    df = df.append(data)
df=df.drop_duplicates('Report Name ')
print(df)
writer = pd.ExcelWriter('pandas_simple.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1')
writer.save()
