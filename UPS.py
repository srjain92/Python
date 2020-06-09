import pyodbc
import pandas as pd
import time
import datetime
import xlsxwriter
import sqlalchemy
import urllib

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=Server Name;'
                      'Database=Database Name;'
                      'UID=abc;'
                      'PWD=xxx;')

script1 =('''
        **********INSERT SQL QUERY***********
        ''')

df1 = pd.read_sql_query(script1, conn)

params = urllib.parse.quote_plus('Driver={SQL Server};'
                                 'Server=Server Name;'
                                 'Database=Database Name;'
                                 'Trusted_Connection=yes;')

conn = sqlalchemy.create_engine('mssql+pyodbc:///?odbc_connect=%s' % params)
connection = conn.connect()
truncate_query = sqlalchemy.text("TRUNCATE TABLE TABLENAME")
connection.execution_options(autocommit=True).execute(truncate_query)

df1.to_sql(name='TABLENAME', con=conn, if_exists='append', index=False, method="multi", chunksize=100)

print('Task Completed')


