import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

def send_report_email(report_file='regression_report.xlsx', plot_file='regression_plot.png', receiver='levi.araujo@gmail.com'):
    # Email configuration (set environment variables or update here)
    sender_email = os.environ.get('EMAIL_SENDER', 'levi.araujo@gmail.com')  # Set EMAIL_SENDER env var
    password = os.environ.get('EMAIL_PASSWORD')  # Set EMAIL_PASSWORD env var

    if not password:
        password = input("Enter your email password (app password): ")

    if not sender_email or not password:
        print("Sender email and password are required.")
        return

    # Create message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver
    msg['Subject'] = 'Regression Analysis Report - 2024 Data'

    body = 'Attached is the latest regression report and chart for 2024 data.'
    msg.attach(MIMEText(body, 'plain'))

    # Attach report
    try:
        with open(report_file, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename= {report_file}")
            msg.attach(part)
    except FileNotFoundError:
        print(f"Report file '{report_file}' not found, skipping attachment.")

    # Attach plot
    try:
        with open(plot_file, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename= {plot_file}")
            msg.attach(part)
    except FileNotFoundError:
        print(f"Plot file '{plot_file}' not found, skipping attachment.")

    # Send email
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, password)
        text = msg.as_string()
        server.sendmail(sender_email, receiver, text)
        server.quit()
        print("Email sent successfully!")
    except Exception as e:
        print(f"Failed to send email: {e}")

if __name__ == "__main__":
    send_report_email()