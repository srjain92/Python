#Code to upload an excel file in the database
import pandas as pd
import xlrd
import pyodbc
import sqlalchemy
import urllib

x1 = pd.ExcelFile(r'\\filename.xlsx')
x1 = x1.sheet_names[0]

df = pd.read_excel(r'\\filename.xlsx', x1)

df.rename(columns={
    'Shipment Tracking Number': 'TrackingNumber'
}, inplace=True)

# df['Shipment Date'] = pd.to_datetime(df['Shipment Date'])
# df['Shipment Delivery Date'] = pd.to_datetime(df['Shipment Delivery Date'])

params = urllib.parse.quote_plus('Driver={SQL Server};'
                                 'Server=Server Name;'
                                 'Database=Database Name;'
                                 'Trusted_Connection=yes;')

conn = sqlalchemy.create_engine('mssql+pyodbc:///?odbc_connect=%s' % params)

df.to_sql(name='File Name', con=conn, if_exists='append', index=False, method="multi", chunksize=100)

print('Task Completed')
