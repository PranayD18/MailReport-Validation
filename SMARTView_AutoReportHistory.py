#Importing Libraries
import os
import pandas as pd
import sqlalchemy
from sqlalchemy import create_engine, text, exc
from sqlalchemy.dialects.mssql import pymssql
import datetime
from datetime import datetime

#MSSQL DB Connection
engine=sqlalchemy.create_engine('mssql+pymssql://SMART:smart@#789@172.16.56.213:1433/ServiceNXT_SMART_Repository')


# List files:
path = os.getcwd()
files = os.listdir(path)
files

#Differentiating '.xlsx' files
files_xlsx = [f for f in files if f[-4:] == 'xlsx']
print("files: ",files_xlsx)

#Defining Dataframe
df = pd.DataFrame()

#Appending All the data from all xlsx files and storing it into df
for f in files_xlsx:
    data = pd.read_excel(f,'Validation')
    df = df.append(data)

print(df)

## Reading MAX of Datetime
df1=pd.DataFrame()
df2=pd.DataFrame()
MaxDatequery='Select MAX([Report DateTime]) FROM [ServiceNXT_SMART_Repository].[dbo].[MailReportAutomation]'
df1['MDate']=pd.read_sql(MaxDatequery,engine)

print(df1)

#print(df['Report DateTime '])
df2=df[df['Report DateTime ']>df1['MDate'][0]]

print(df2)

#Inserting df data into table
df2.to_sql(
      name='MailReportAutomation',
      con=engine,
      index=False,
      if_exists='append')

print("Data Updated Successfully")

#df2['LastModifiedDate']=pd.to_datetime('today').replace(microsecond=0)
#print(df2)
