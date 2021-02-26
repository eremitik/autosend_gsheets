import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import smtplib
import mimetypes
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.text import MIMEText

print('Starting script...')
#get data
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('creds_pub.json', scope)
client = gspread.authorize(creds)
sheet = client.open('SHEET_NAME').worksheet('WORKSHEET_NAME')

data = sheet.get_all_records()
row = sheet.row_values('ENTER_NUMBER')
col = sheet.col_values('ENTER_NUMBER')
cell = sheet.cell('ENTER_CELL').value
vol_data = sheet.get('ENTER_RANGE')

df = pd.DataFrame(vol_data)
csv_data = df.to_csv(header=False, index=False)
print(csv_data)


#saving to excel file
df.to_excel('NAME.xlsx', sheet_name='SHEET_NAME', index=False, header=False)
print('Saved to excel file.')


#sending email
emailfrom = "YOUR_EMAIL"
emailto = "RECEIPIENTS_HERE"
fileToSend = "SHEET_NAME.xlsx"
username = "YOUR_EMAIL"
password = "PASSCODE"

msg = MIMEMultipart()
msg["From"] = emailfrom
msg["To"] = ",".join(emailto)
msg["Subject"] = "SUBJECT"
msg.preamble = "BODY_MESSAGE"

ctype, encoding = mimetypes.guess_type(fileToSend)
if ctype is None or encoding is not None:
    ctype = "application/octet-stream"

maintype, subtype = ctype.split("/", 1)

if maintype == "text":
    fp = open(fileToSend)
    # Note: we should handle calculating the charset
    attachment = MIMEText(fp.read(), _subtype=subtype)
    fp.close()
elif maintype == "image":
    fp = open(fileToSend, "rb")
    attachment = MIMEImage(fp.read(), _subtype=subtype)
    fp.close()
elif maintype == "audio":
    fp = open(fileToSend, "rb")
    attachment = MIMEAudio(fp.read(), _subtype=subtype)
    fp.close()
else:
    fp = open(fileToSend, "rb")
    attachment = MIMEBase(maintype, subtype)
    attachment.set_payload(fp.read())
    fp.close()
    encoders.encode_base64(attachment)
attachment.add_header("Content-Disposition", "attachment", filename=fileToSend)
msg.attach(attachment)

server = smtplib.SMTP("smtp.gmail.com:587")
server.starttls()
server.login(username,password)
server.sendmail(emailfrom, emailto, msg.as_string())
server.quit()
print('Email has been sent.')
