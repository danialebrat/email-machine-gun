import smtplib
import mimetypes
from email.message import EmailMessage
import pandas as pd
from bs4 import BeautifulSoup
import re
import time

# Path to your List (emails, professors, and , etc)
csv_name = 'list.csv'
df_professors = pd.read_csv(csv_name)

# IMPORTANT SUGGESTIONS!!!
# Consider adding a wait time to control the number of email per hour
# consider at least 5 Min before sending the next email to avoid your email getting blocked
# Don't send more than 100 emails per day or 20 emails per hour
# you can control all of this using a simple IF condition



for index, row in df_professors.iterrows():
    send_status = row['send_status']
    if send_status != 0:
        continue
    message = EmailMessage()
    sender = "Put your Email here"
    message['From'] = sender
    message['To'] = row['email']

    if row['subject'] == "none":
        message['Subject'] = "Prospective Ph.D. Student"
    else:
        message['Subject'] = str(row['subject'])

    if row['field'] == "music":
        with open("template_music.html", "r", encoding='utf-8') as f:
            html = f.read()
            soup = BeautifulSoup(html, "html.parser")
            target = soup.find_all(text=re.compile(r'{name}'))
            for v in target:
                v.replace_with(v.replace('{name}', row['name']))

            target = soup.find_all(text=re.compile(r'{uni}'))
            for v in target:
                v.replace_with(v.replace('{uni}', row['uni']))

            target = soup.find_all(text=re.compile(r'{group}'))
            for v in target:
                if row['group'] == "none":
                    temp = ""
                    v.replace_with(v.replace('{group}', str(temp)))
                else:
                    temp = ' ' + 'at ' + str(row['group'])
                    v.replace_with(v.replace('{group}', temp))


            message.set_content(str(soup), subtype='html')
    else:
        with open("template.html", "r", encoding='utf-8') as f:
            html = f.read()
            soup = BeautifulSoup(html, "html.parser")
            target = soup.find_all(text=re.compile(r'{name}'))
            for v in target:
                v.replace_with(v.replace('{name}', row['name']))

            target = soup.find_all(text=re.compile(r'{uni}'))
            for v in target:
                v.replace_with(v.replace('{uni}', row['uni']))

            target = soup.find_all(text=re.compile(r'{group}'))
            for v in target:
                if row['group'] == "none":
                    temp = ""
                    v.replace_with(v.replace('{group}', str(temp)))
                else:
                    temp = ' ' + 'at ' + str(row['group'])
                    v.replace_with(v.replace('{group}', temp))

            message.set_content(str(soup), subtype='html')


    with open('your_CV.pdf', 'rb') as file:
        mime_type, _ = mimetypes.guess_type('your_CV.pdf')
        mime_type, mime_subtype = mime_type.split('/')
        message.add_attachment(file.read(),
                               maintype=mime_type,
                               subtype=mime_subtype,
                               filename='your_CV.pdf')

    if row['transcript'] == True:

        with open('your_transcript.pdf', 'rb') as file:
            mime_type, _ = mimetypes.guess_type('your_transcript.pdf')
            mime_type, mime_subtype = mime_type.split('/')
            message.add_attachment(file.read(),
                                   maintype=mime_type,
                                   subtype=mime_subtype,
                                   filename='your_transcript.pdf')

        with open('your_transcript.pdf', 'rb') as file:
            mime_type, _ = mimetypes.guess_type('your_transcript.pdf')
            mime_type, mime_subtype = mime_type.split('/')
            message.add_attachment(file.read(),
                                   maintype=mime_type,
                                   subtype=mime_subtype,
                                   filename='your_transcript.pdf')


    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login("your_email@gmail.com", 'your_password(Google Generated Code)')
    mail_server.send_message(message)
    print("successful send email to ", str(message['To']))
    df_professors.loc[index, 'send_status'] = 1
    df_professors.to_csv(csv_name, index=False)
    mail_server.quit()

    # waiting for 5 min
    time.sleep(300)
