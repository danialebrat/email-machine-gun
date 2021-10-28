from dotenv import load_dotenv
from email.message import EmailMessage
from bs4 import BeautifulSoup
import smtplib
import mimetypes
import pandas as pd
import re
import os

load_dotenv()

GMAIL_USER = os.getenv('GMAIL_USER')
GMAIL_PASS = os.getenv('GMAIL_PASS')
GMAIL_SENDER = os.getenv('GMAIL_SENDER')
GMAIL_SUBJECT = os.getenv('GMAIL_SUBJECT')
MY_NAME = os.getenv('MY_NAME')
PROFESSORS_CSV_FILE = 'professor_list.csv'


df_professors = pd.read_csv(PROFESSORS_CSV_FILE)
for index, row in df_professors.iterrows():
    send_status = row['send_status']
    if send_status != 0:
        continue
    message = EmailMessage()
    message['From'] = GMAIL_SENDER
    message['Subject'] = GMAIL_SUBJECT
    message['To'] = row['email']

    with open("template.html", "r", encoding='utf-8') as f:
        html = f.read()
        soup = BeautifulSoup(html, "html.parser")
        target = soup.find_all(text=re.compile(r'{prof_name}'))
        for v in target:
            v.replace_with(v.replace('{prof_name}', row['name']))

        target = soup.find_all(text=re.compile(r'{my_name}'))
        for v in target:
            v.replace_with(v.replace('{my_name}', MY_NAME))

        message.set_content(str(soup), subtype='html')
        target = soup.find_all(text=re.compile(r'{university}'))
        for v in target:
            v.replace_with(v.replace('{university}', row['university']))

        message.set_content(str(soup), subtype='html')

    with open('CV.pdf', 'rb') as file:
        mime_type, _ = mimetypes.guess_type('CV.pdf')
        mime_type, mime_subtype = mime_type.split('/')
        message.add_attachment(file.read(),
                               maintype=mime_type,
                               subtype=mime_subtype,
                               filename='CV.pdf')

    with open('Transcript.pdf', 'rb') as file:
        mime_type, _ = mimetypes.guess_type('Transcript.pdf')
        mime_type, mime_subtype = mime_type.split('/')
        message.add_attachment(file.read(),
                               maintype=mime_type,
                               subtype=mime_subtype,
                               filename='Transcript.pdf')

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(GMAIL_USER, GMAIL_PASS)
    mail_server.send_message(message)
    print("successful send email to ", str(message['To']))
    df_professors.loc[index, 'send_status'] = 1
    df_professors.to_csv(PROFESSORS_CSV_FILE, index=False)
    mail_server.quit()
