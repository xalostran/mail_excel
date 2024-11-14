import os
import imaplib
import email
import dotenv
import csv
import pandas as pd

dotenv.load_dotenv()

username = os.getenv('EMAIL')
password = os.getenv('PASSWORD')
imap_host = os.getenv('HOST')
imap_port = os.getenv('PORT')

with open('name_of_.csv', mode='w', newline='', encoding='utf-8') as csv_file:
    fieldnames = ['Subject', 'From', 'To', 'Date', 'Content']
    writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
    writer.writeheader()
    iMap = imaplib.IMAP4_SSL(imap_host, imap_port)
    iMap.login(username, password)

    iMap.select("Inbox")
    is_ok, msg = iMap.search(None, "ALL")

    if is_ok:
        print("OK", msg)

    for msgs in msg[0].split():
        is_ok, data = iMap.fetch(msgs, '(RFC822)')
        message = email.message_from_bytes(data[0][1])

        print(message.get('Subject'))
        print(message.get('From'))
        print(message.get('To'))
        print(message.get('Date'))

        subject = message.get('Subject')
        from_ = message.get('From')
        to = message.get('To')
        date = message.get('Date')
        content = ""

        for part in message.walk():
            if part.get_content_type() == "text/plain":
                content += part.get_payload(decode=True).decode(
                    part.get_content_charset(), errors='replace')
                # decodear från binärt format till textform

        writer.writerow({  # Ungefär som createcell i java
            'Subject': subject,
            'From': from_,
            'To': to,
            'Date': date,
            'Content': content
        })


iMap.close()
