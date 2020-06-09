import pyodbc
import pandas as pd
import time
import datetime
import xlsxwriter
import sqlalchemy
import urllib

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=dds-services-uscen01.database.windows.net;'
                      'Database=DDSLab_Services;'
                      'UID=ddsreporting;'
                      'PWD=repor62937!g;')

script1 =('''SELECT [ID]
            ,[DDS_CaseID] AS [CaseID]
            ,CONVERT(DATE,[DDS_ShipDate]) AS [Ship Date]
            FROM [DDSLab_Services].[dbo].[ShippingEndofDay]
            WHERE [Deleted] = 0''')

df1 = pd.read_sql_query(script1, conn)

params = urllib.parse.quote_plus('Driver={SQL Server};'
                                 'Server=dds-sql03;'
                                 'Database=DLPLUS;'
                                 'Trusted_Connection=yes;')

conn = sqlalchemy.create_engine('mssql+pyodbc:///?odbc_connect=%s' % params)
connection = conn.connect()
truncate_query = sqlalchemy.text("TRUNCATE TABLE DDSNet_UPS_IPD")
connection.execution_options(autocommit=True).execute(truncate_query)

df1.to_sql(name='DDSNet_UPS_IPD', con=conn, if_exists='append', index=False, method="multi", chunksize=100)

print('Task Completed')


