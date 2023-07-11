import xlsxwriter
import os
import pandas as pd
import cx_Oracle
import openpyxl
from pandas import ExcelWriter
from openpyxl import Workbook

from email.message import EmailMessage
from email.utils import make_msgid
import mimetypes
import smtplib

# Oracle database connection
conn = cx_Oracle.connect("username", "password", "db name")
print("Oracle database connected")

# POPRRS queryset 

df1=pd.read_sql("""query
""",con=conn)

# Removing excel excels seperatly
os.remove("Report.xlsx")
print("Excel removed")

# Queryset convert to seperate excels and save 

writer=pd.ExcelWriter("Report.xlsx",engine="openpyxl")
df1.to_excel(writer,sheet_name="Sheet1",index=False)

writer.save()
print("Excels Downloaded")

s = smtplib.SMTP(host='smtp.office365.com', port=587)
s.starttls()

s.login('email address', 'password')
msg = EmailMessage()

print("Ready for mailing")

# generic email headers
msg['Subject'] = 'Report subject'

msg['From'] = 'Name <email address>'
msg['To'] =  'Name <email address>'

with open(r"C:\coutlook_automation\Report.xlsx", 'rb') as ra:
    attachment = ra.read()
msg.add_related(attachment, maintype='application', subtype='xlsx', filename='Report.xlsx')

msg.add_alternative("""\
    <html>
        <body>
            <p><i>Dear Team,</i></p>
            <p> </p>
            <p><i> Please find the attachments.</i></p>
            <p></p>
            <p> <i>Thanks & Regards,<br>
                ( This is an autogenerated mail )<br>
        Department Name <br> 
            </i></p>
        </body>
    </html>
    """ ,subtype='html')


s.send_message(msg)

print("Mail sent")

conn.close()




