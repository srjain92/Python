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

from xlsxwriter import worksheet

today = datetime.date.today()
first = today.replace(day=1)
lastMonth = first - datetime.timedelta(days=1)
lastMonth = (lastMonth.strftime("%B"))

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=Server Name;'
                      'Database=Database name;'
                      'Trusted_Connection=yes;')

script1 = ('''
           ******INSERT SQL QUERY******   
            ''')

script2 = ('''
           ******INSERT SQL QUERY******     
                ''')

script3 = ('''
           ******INSERT SQL QUERY******     
                ''')

### Create a Pandas dataframe from the data
df1 = pd.read_sql_query(script1, conn)
df2 = pd.read_sql_query(script2, conn)
df3 = pd.read_sql_query(script3, conn)

filename = 'ABC' + '_' + lastMonth + '.xlsx'
sheetname1 = 'ABC'
sheetname2 = 'ABC'
sheetname3 = 'ABC'

### Create a Pandas Excel writer using XlsxWriter as the engine
writer = pd.ExcelWriter(filename, engine='xlsxwriter')

### Convert the dataframe to an XlsxWriter Excel object
df1.to_excel(writer, sheet_name=sheetname1, index=False)
df2.to_excel(writer, sheet_name=sheetname2, index=False)
df3.to_excel(writer, sheet_name=sheetname3, index=False)

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
cellformat = workbook.add_format({'num_format': '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'})
worksheet = writer.sheets[sheetname1]
worksheet.set_column('H:H', cell_format=cellformat)
worksheet.set_column('J:J', cell_format=cellformat)
auto_width_column(df1)

worksheet = writer.sheets[sheetname2]
worksheet.set_column('M:M', cell_format=cellformat)
auto_width_column(df2)

worksheet = writer.sheets[sheetname3]
auto_width_column(df3)

### Save the Excel Workbook
writer.save()

### Email sending code
email_user = 'abc@gmail.com'
email_password = 'xxx'
email_list = ['abc@gmail.com']
email_send = ', '.join(email_list)

subject = 'ABC ' + lastMonth

msg = MIMEMultipart()
msg['From'] = email_user
msg['To'] = email_send
msg['Subject'] = subject

body = '''Hey 

Thank you'''.format(lastMonth=lastMonth)

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
