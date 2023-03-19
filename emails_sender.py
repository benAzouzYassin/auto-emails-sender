from email import encoders
from email.mime.base import MIMEBase
from pathlib import Path
import pandas as pd
from time import  sleep
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import datetime

print("\n \nIMPORTANT NOTES !! :" )
print("    - 2 steps authentication should be active in the email of the sender!")
print("    - Enter the application password of the sender email not the personal one!!!")
print("    - the excel file data.xlsx should exist at the same folder contains this file!")
sender_email = input("sender email : ")
sender_password= input("sender application password : ")

while True:
    excel_file = "data.xlsx"
    df = pd.read_excel(excel_file, sheet_name=0)
    emails = list(df.to_dict()["email"].values())
    send_date= str(list(df.to_dict()["send date"].values())[0])
    subject = list(df.to_dict()["subject"].values())[0]
    content = list(df.to_dict()["content"].values())[0]
    attachments = list(df.to_dict()["attachments"].values())

    todayDate = datetime.date.today().strftime('%d/%m/%Y')


    if (send_date ==todayDate):
        for mail in emails :
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['Subject'] = subject
            msg['To'] = mail
            msg.attach(MIMEText(content, "plain"))
            for fileName in attachments:
                path = Path(str(fileName))
                try:
                    with path.open("rb") as f:
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(f.read())
                        encoders.encode_base64(part)
                        part.add_header(
                            "Content-Disposition",
                            f"attachment; filename={path.name}",
                        )
                        msg.attach(part)
                except:
                    pass
            try:
                server = smtplib.SMTP(host='smtp.gmail.com', port=587)
                server.ehlo()
                server.starttls()
                server.login(sender_email,sender_password)
                server.send_message(msg)
                print(f"sended to {mail} in : {todayDate} ")
            except:
                print("error on sending the mail")
    sleep(86400)
