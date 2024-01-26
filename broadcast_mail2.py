import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
import glob
import os

def generate_email_bodies(team_list):
    email_bodies = []
    for team in team_list:
        body = f"""
<h1>Header 1</h1>
<p>Paragraph</p>
        """
        email_bodies.append(body)
    return email_bodies

def setup_smtp_server():
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    return server

def login_to_email(server, mail_account, password):
    try:
        server.login(mail_account, password)
    except smtplib.SMTPAuthenticationError:
        print("Failed to login")
        return False
    return True

# TAMBAH FILE ATTACHMENT
def add_attachments(msg, team):
    
    dir = "attachment file directory"
    filenames = glob.glob(f"{dir}{team}*")
    for filename in filenames:
        with open(filename, "rb") as binary_file:
            payload = MIMEBase("application", "octet-stream")
            payload.set_payload(binary_file.read())
            encoders.encode_base64(payload)
            base_filename = os.path.basename(filename)
            payload.add_header("Content-Disposition", "attachment", filename=base_filename)
            msg.attach(payload)
    return msg

def send_emails(server, email_bodies, mail_account, mail_to, team_list, sender_name):
    for i in range(len(team_list)):
        msg = MIMEMultipart()
        msg['From'] = sender_name
        msg['To'] = mail_to[i]
        msg['Subject'] = "YOUR SUBJECT"
        msg.attach(MIMEText(email_bodies[i], 'html'))

        msg = add_attachments(msg, team_list[i])

        text = msg.as_string()
        try:
            server.sendmail(mail_account, mail_to[i], text)
            print(f"Mail sent to {mail_to[i]} [{team_list[i]}]")
        except smtplib.SMTPException as e:
            print(f"Failed to send mail to {mail_to[i]} [{team_list[i]}]. Error: {e}")


def read_excel_file(file_path):
    xl = pd.read_excel(file_path, sheet_name='Sheet1')
    mail_to = xl['Email'].tolist()
    team_list = xl['NamaTim'].tolist()
    return mail_to, team_list

if __name__ == "__main__":
    # get information from excel file
    mail_to, team_list = read_excel_file('./coba.xlsx')

    sender_name = "TECHCOMFEST 2024"
    # use google app password if you are using gmail
    # different email provider may have different way to generate app password
    mail_account = 'your@email.com'
    password = 'yourpass' 

    email_bodies = generate_email_bodies(team_list)
    server = setup_smtp_server()

    if login_to_email(server, mail_account, password):
        send_emails(server, email_bodies, mail_account, mail_to, team_list, sender_name)