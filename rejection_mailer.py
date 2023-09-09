import os
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Fetch SMTP credentials from environment variables
smtp_server = os.environ.get('SMTP_SERVER')
smtp_port = os.environ.get('SMTP_PORT')
smtp_username = os.environ.get('SMTP_USERNAME')
smtp_password = os.environ.get('SMTP_PASSWORD')

# Function to read Excel file and return DataFrame
def read_applicants_from_excel(filepath):
    try:
        return pd.read_excel(filepath)
    except Exception as e:
        print(f"An error occurred while reading the Excel file: {e}")
        return None

# Function to send email
def send_email(email_address, name):
    try:
        subject = "Your Application Status"
        body = f"Dear {name},\n\nThank you for applying. Unfortunately, we cannot move forward with your application at this time.\n\nBest,\nCompany XYZ"
        
        msg = MIMEMultipart()
        msg["From"] = smtp_username
        msg["To"] = email_address
        msg["Subject"] = subject
        
        msg.attach(MIMEText(body, 'plain'))
        
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_username, smtp_password)
        
        text = msg.as_string()
        server.sendmail(smtp_username, email_address, text)
        server.quit()
        
        print(f"Email sent to {email_address}")
        
    except smtplib.SMTPAuthenticationError:
        print("SMTP Authentication Error: The username or password you entered is incorrect.")
    except smtplib.SMTPConnectError:
        print(f"Could not connect to the server: {smtp_server}")
    except Exception as e:
        print(f"An error occurred while sending the email: {e}")

if __name__ == "__main__":
    # Read applicant data from Excel file
    df = read_applicants_from_excel("path/to/applicants.xlsx")
    
    if df is not None:
        # Loop through each row in the DataFrame
        for index, row in df.iterrows():
            email = row['Email']
            name = row['Name']
            
            # Send email to each applicant
            send_email(email, name)