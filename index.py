import os
import time
import smtplib
import pandas as pd
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.message import EmailMessage
from dotenv import load_dotenv  # To manage sensitive data

# Load environment variables
load_dotenv()
SMTP_USERNAME = os.getenv('SMTP_USERNAME')
SMTP_PASSWORD = os.getenv('SMTP_PASSWORD')

def send_attachment_email(subject, body, to_email, file_path):
    """Send an email with an attachment."""
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = SMTP_USERNAME
    msg['To'] = to_email
    msg.set_content(body)

    # Attach the file
    file_name = os.path.basename(file_path)
    with open(file_path, 'rb') as file:
        msg.add_attachment(file.read(), maintype='application', subtype='octet-stream', filename=file_name)

    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(SMTP_USERNAME, SMTP_PASSWORD)
        server.send_message(msg)
        print(f'Email sent to {to_email}!')

def send_reminder_email(to_address, cc_address, subject, body_content, color):
    """Send reminder email with color-coded message."""
    html_body = f"""
    <html>
    <body style="font-family: Tahoma, sans-serif;">
        <p>Dear {body_content['employee_name']},</p>
        <p>Your <span style="background-color: {color};">
        <b>"{body_content['category']}"</b></span> training expires on <b>{body_content['expiry_date']}</b>.
        Please renew it in time to avoid disruptions.</p>
    </body>
    </html>
    """

    msg = MIMEText(html_body, 'html')
    msg['Subject'] = subject
    msg['From'] = SMTP_USERNAME
    msg['To'] = to_address
    msg['Cc'] = cc_address

    all_recipients = to_address.split(",") + cc_address.split(",")

    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(SMTP_USERNAME, SMTP_PASSWORD)
        server.sendmail(SMTP_USERNAME, all_recipients, msg.as_string())

def load_and_categorize_data(file_path):
    """Load Excel files and categorize employees by expiry dates."""
    df = pd.read_excel(file_path, header=[0, 1])
    today = datetime.today().date()
    categories = ['Oil and gas services', 'Lifting equipment', 'Hazardous transport']

    colors = {'expired': 'red', 'one_week': 'orange', 'two_weeks': 'yellow', 'one_month': 'green'}

    categorized_data = {cat: {'expired': [], 'one_week': [], 'two_weeks': [], 'one_month': []} for cat in categories}

    for _, row in df.iterrows():
        for category in categories:
            expiry_date = row[('Expiry Date', category)].date()
            days_to_expiry = (expiry_date - today).days

            if days_to_expiry < 0:
                key = 'expired'
            elif days_to_expiry <= 7:
                key = 'one_week'
            elif days_to_expiry <= 14:
                key = 'two_weeks'
            else:
                key = 'one_month'

            categorized_data[category][key].append(row)

    return categorized_data

def main():
    file_path = 'WCM_D&M_Safety Instructions.xlsx'
    categorized_data = load_and_categorize_data(file_path)

    for category, data in categorized_data.items():
        for key, employees in data.items():
            for employee in employees:
                body_content = {
                    'employee_name': employee[('Employee', 'Unnamed: 2_level_1')],
                    'category': category,
                    'expiry_date': employee[('Expiry Date', category)].strftime('%B %d, %Y')
                }
                send_reminder_email(
                    employee[('Email', 'Unnamed: 3_level_1')],
                    employee[('Manager Email', 'Unnamed: 10_level_1')],
                    f'{category} Training Expiration Notice',
                    body_content,
                    colors[key]
                )

if __name__ == '__main__':
    main()
