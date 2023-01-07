import pandas as pd
import pdfkit 
import glob 
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
#Setting Up system config files for the pakage
path_wkhtmltopdf = 'C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe'
config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
#Getting all filepaths for the directory
filepaths=glob.glob("Invoices/*.xlsx")
frames=[]
for file in filepaths:
    data=pd.read_excel(file)
    frames.append(data)
#Creating a final dataset that will have all the Excel files
final_data=pd.DataFrame(pd.concat(frames))
total_sales=round(sum(final_data.total_price),0)
print("Your final dataset is created.")
#Saving The dataset as a html file
final_data.to_html('input.html')
#Coverting the file as a pdf file
pdfkit.from_file('input.html', 'output.pdf',configuration=config)
print("Your PDF file is created.")

body = f"Hello , Here is Your Daily Report Today you have recorded a sales of ${total_sales}"
# put your email here
sender = 'meenakshi.bhtt@gmail.com'
# get the password in the gmail (manage your google account, click on the avatar on the right)
# then go to security (right) and app password (center)
# insert the password and then choose mail and this computer and then generate
# copy the password generated here
password = 'ztyuzvfbrxqawcbv'
# put the email of the receiver here
receiver = 'ab0358031@gmail.com'

#Setup the MIME
message = MIMEMultipart()
message['From'] = sender
message['To'] = receiver
message['Subject'] = 'This email has an attachment, a pdf file'

message.attach(MIMEText(body, 'plain'))

pdfname = 'output.pdf'

# open the file in bynary
binary_pdf = open(pdfname, 'rb')

payload = MIMEBase('application', 'octate-stream', Name=pdfname)
# payload = MIMEBase('application', 'pdf', Name=pdfname)
payload.set_payload((binary_pdf).read())

# enconding the binary into base64
encoders.encode_base64(payload)

# add header with pdf name
payload.add_header('Content-Decomposition', 'attachment', filename=pdfname)
message.attach(payload)

#use gmail with port
session = smtplib.SMTP('smtp.gmail.com', 587)

#enable security
session.starttls()

#login with mail_id and password
session.login(sender, password)

text = message.as_string()
session.sendmail(sender, receiver, text)
session.quit()
print('Mail Sent')

