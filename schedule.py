#schedual for sending
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import schedule
import time

# Define the function to be scheduled
def job():
    subject = "Test Email"
    body = "This is a test email sent from a Python script!"
    to_email = "phonsreyvang89@gmail.com"  # Replace with recipient's email
    send_email(subject, body, to_email)

# Schedule the job
schedule.every().day.at("08:00").do(job)  # Schedule for 8:00 AM daily

# Keep the script running
while True:
    schedule.run_pending()  # Check if a scheduled task is pending to run
    time.sleep(60)  # Wait for one minute

