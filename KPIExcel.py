import pyodbc
import pandas as pd
import xlsxwriter
import time
import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

today = datetime.date.today()
first = today.replace(day=1)
lastMonth = first - datetime.timedelta(days=1)
lastMonth = (lastMonth.strftime("%m_%Y"))

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=Server Name;'
                      'Database=Database name;'
                      'Trusted_Connection=yes;')

script1 = ('''
           *******INSERT SQL QUERY*******
          ''')

script2_1 = ('''
           *******INSERT SQL QUERY*******
          ''')

script2_2 = ('''
           *******INSERT SQL QUERY*******
          ''')

script2_3 = ('''
           *******INSERT SQL QUERY*******
          ''')

script2_4 = ('''
           *******INSERT SQL QUERY*******
          ''')

script2_5 = ('''
           *******INSERT SQL QUERY*******
          ''')

script2_6 = ('''
           *******INSERT SQL QUERY*******
          ''')

script2_7 = ('''
           *******INSERT SQL QUERY*******
          ''')

script3 = ('''
           *******INSERT SQL QUERY*******
          ''')

script4 = ('''
           *******INSERT SQL QUERY*******
          ''')

script5 = ('''
           *******INSERT SQL QUERY*******
          ''')

script6 = ('''
           *******INSERT SQL QUERY*******
          ''')

### Create a Pandas dataframe from the data
df1 = pd.read_sql_query(script1, conn)
df2_1 = pd.read_sql_query(script2_1, conn)
df2_2 = pd.read_sql_query(script2_2, conn)
df2_3 = pd.read_sql_query(script2_3, conn)
df2_4 = pd.read_sql_query(script2_4, conn)
df2_5 = pd.read_sql_query(script2_5, conn)
df2_6 = pd.read_sql_query(script2_6, conn)
df2_7 = pd.read_sql_query(script2_7, conn)
df3 = pd.read_sql_query(script3, conn)
df4 = pd.read_sql_query(script4, conn)
df5 = pd.read_sql_query(script5, conn)
df6 = pd.read_sql_query(script6, conn)

filename = 'KPI_Monthly' + '_' + lastMonth + '.xlsx'
sheetname1 = 'ABC' + '_' + lastMonth
sheetname2 = 'ABC' + '_' + lastMonth
sheetname3 = 'ABC' + '_' + lastMonth
sheetname4 = 'ABC' + '_' + lastMonth
sheetname5 = 'ABC' + '_' + lastMonth
sheetname6 = 'ABC' + '_' + lastMonth

### Create a Pandas Excel writer using XlsxWriter as the engine
writer = pd.ExcelWriter(filename, engine='xlsxwriter')

### Convert the dataframe to an XlsxWriter Excel object
df1.to_excel(writer, sheet_name=sheetname1, index=False)
df2_1.to_excel(writer, sheet_name=sheetname2, index=False)
df2_2.to_excel(writer, sheet_name=sheetname2, index=False, startrow=4)
df2_3.to_excel(writer, sheet_name=sheetname2, index=False, startrow=17)
df2_4.to_excel(writer, sheet_name=sheetname2, index=False, startrow=32)
df2_5.to_excel(writer, sheet_name=sheetname2, index=False, startrow=39)
df2_6.to_excel(writer, sheet_name=sheetname2, index=False, startrow=46)
df2_7.to_excel(writer, sheet_name=sheetname2, index=False, startrow=53)
df3.to_excel(writer, sheet_name=sheetname3, index=False)
df4.to_excel(writer, sheet_name=sheetname4, index=False)
df5.to_excel(writer, sheet_name=sheetname5, index=False)
df6.to_excel(writer, sheet_name=sheetname6, index=False)

# create function for excel column auto width
def auto_width_column(df):
    for i, col in enumerate(df.columns):
        # find length of column i
        column_len = df[col].astype(str).str.len().max()
        # Setting the length if the column header is larger than the max column value length
        column_len = max(column_len, len(col)) + 2
        worksheet.set_column(i, i, column_len)

### Excel formatting
workbook = writer.book
worksheet = writer.sheets[sheetname1]
auto_width_column(df1)

worksheet = writer.sheets[sheetname2]
auto_width_column(df2_1)
auto_width_column(df2_3)
cellformat = workbook.add_format({'bold': True, 'font_color': 'red', 'bg_color': 'yellow'})
worksheet.write_string(38, 0, 'FIXED', cell_format=cellformat)
worksheet.write_string(45, 0, 'REMOVABLE', cell_format=cellformat)
worksheet.write_string(52, 0, 'ORTHO', cell_format=cellformat)

worksheet = writer.sheets[sheetname3]
auto_width_column(df3)
worksheet = writer.sheets[sheetname4]
auto_width_column(df4)
worksheet = writer.sheets[sheetname5]
auto_width_column(df5)
worksheet = writer.sheets[sheetname6]
auto_width_column(df6)

### Save the Excel Workbook
writer.save()

### Email sending code
email_user = 'abc@gmail.com'
email_password = 'xxx'
email_list = ['abc@gmail.com']
email_send = ', '.join(email_list)

subject = 'KPI Monthly Data ' + lastMonth

msg = MIMEMultipart()
msg['From'] = email_user
msg['To'] = email_send
msg['Subject'] = subject

body = '''Hey 
         
Please find attached'''

msg.attach(MIMEText(body, 'plain'))

file_name = filename
attachment = open(file_name, 'rb')

part = MIMEBase('application', 'octet-stream')
part.set_payload(attachment.read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', "attachment; filename= " + file_name)

msg.attach(part)
text = msg.as_string()
server = smtplib.SMTP('smtp-mail.outlook.com', 587)
server.starttls()
server.login(email_user, email_password)

server.sendmail(email_user, email_list, text)
server.quit()

print('Task Completed')
