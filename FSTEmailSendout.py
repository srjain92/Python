import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

filename = r'\\dds-file04\dds lab\Salesforce Issues Log\FST Detail Information Data Extract.xlsx'
# filename = 'S:\\Salesforce Issues Log\FST Detail Information Data Extract.xlsx'

# Email sending code
email_user = 'saurabh.jain@ddslab.com'
email_password = 'J@in2018oct'
email_list = ['Rose.Vazquez@ddslab.com', 'Kent.Billiter@ddslab.com', 'lacey.mitchell@ddslab.com', 'Petar.Andjelic@ddslab.com',
              'Marge.Nicholson@ddslab.com', 'Michael.Radke@ddslab.com', 'Katrina.Dahdah@ddslab.com', 'Preston.Bryan@ddslab.com',
              'Nicolle.Makle@ddslab.com', 'Michelle.Brockman@ddslab.com', 'Brianna.Voshell@ddslab.com', 'Terrance.Hopson@ddslab.com'
              , 'Terry.Davis@ddslab.com', 'Angela.Warren@ddslab.com', 'Brittany.Diehl@ddslab.com', 'saurabh.jain@ddslab.com']
email_send = ', '.join(email_list)

subject = 'FST Underlying data behind Tableau dashboard'

msg = MIMEMultipart()
msg['From'] = email_user
msg['To'] = email_send
msg['Subject'] = subject

body = '''Hey Team,

Please find attached the FST scanner training data

This is an automated email. For any questions regarding the data, please reach out to Rose.

Thank you'''

msg.attach(MIMEText(body, 'plain'))

file_name = filename
attachment = open(file_name, 'rb')

part = MIMEBase('application', 'octet-stream')
part.set_payload(attachment.read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', "attachment; filename= " + "FST Detail Information Data Extract.xlsx")

msg.attach(part)
text = msg.as_string()
server = smtplib.SMTP('smtp-mail.outlook.com', 587)
server.starttls()
server.login(email_user, email_password)

server.sendmail(email_user, email_list, text)
server.quit()

print('Task Completed')
