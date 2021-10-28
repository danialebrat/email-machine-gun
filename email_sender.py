import smtplib
import mimetypes
from email.message import EmailMessage
import pandas as pd
from bs4 import BeautifulSoup
import re


email = "Your_Email@gmail.com"
password = "Put Your Password Here"

# The location of your master list
csv_name = 'Your_Master_List.csv'
df_professors = pd.read_csv(csv_name)

# Put your Files Here
cv = "YourCV.pdf"
transcript = "Your_Transcript.pdf"
template = "template.html"


# The subject of Your email
subject = "Subject of your email"


for index, row in df_professors.iterrows():

    send_status = row['send_status']
    if send_status != 0:
        continue
    message = EmailMessage()
    sender = email
    message['From'] = sender
    message['To'] = row['email']
    message['Subject'] = subject

    with open(template, "r", encoding='utf-8') as f:
        html = f.read()
        soup = BeautifulSoup(html, "html.parser")
        target = soup.find_all(text=re.compile(r'{name}'))
        for v in target:
            v.replace_with(v.replace('{name}', row['name']))

        target = soup.find_all(text=re.compile(r'{university}'))
        for v in target:
            v.replace_with(v.replace('{university}', row['university']))

        message.set_content(str(soup), subtype='html')

    with open(cv, 'rb') as file:
        mime_type, _ = mimetypes.guess_type(cv)
        mime_type, mime_subtype = mime_type.split('/')
        message.add_attachment(file.read(),
                               maintype=mime_type,
                               subtype=mime_subtype,
                               filename=cv)

    with open(transcript, 'rb') as file:
        mime_type, _ = mimetypes.guess_type(transcript)
        mime_type, mime_subtype = mime_type.split('/')
        message.add_attachment(file.read(),
                               maintype=mime_type,
                               subtype=mime_subtype,
                               filename=transcript)

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(email, password)
    mail_server.send_message(message)
    print("successful send email to ", str(message['To']))
    df_professors.loc[index, 'send_status'] = 1
    df_professors.to_csv(csv_name, index=False)
    mail_server.quit()
