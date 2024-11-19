import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os  # Import os to use environment variables
import schedule
import time
from openpyxl import load_workbook  # Import to handle Excel files

# Load Excel File (Make sure the file path is correct)
excel_file = "Book2.xlsx"  # Replace with your Excel file path
workbook = load_workbook(excel_file)
sheet = workbook.active

def send_email():
    # Use environment variables for email and password to avoid exposing them in the code
    from_email = "soklen.nak@student.passerellesnumeriques.org"  # Sender's email
    to_email = "soklennak@gmail.com"  # Recipient's email
    subject = "Automated Scheduled Email"
    body = "This is an automatically sent email by len."

    # Get the Gmail app password from environment variable (make sure it's set)
    app_password = os.getenv("GMAIL_APP_PASSWORD")  # Set this as an environment variable

    if app_password is None:
        print("Error: GMAIL_APP_PASSWORD environment variable is not set.")
        return

    # Create the email content
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # Set up the SMTP server and send the email
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)  # Gmail's SMTP server and port
        server.starttls()  # Secure the connection
        server.login(from_email, app_password)  # Use the app password
        text = msg.as_string()
        server.sendmail(from_email, to_email, text)
        print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Failed to send email: {e}")
    finally:
        server.quit()

def job():
    print("I'm working...")
    send_email()

# Schedule jobs to run at various times
schedule.every(2).minutes.do(job)  # Send email every 2 minutes
schedule.every().hour.do(job)  # Send email every hour
schedule.every().day.at("10:30").do(job)  # Send email daily at 10:30 AM
schedule.every().monday.do(job)  # Send email every Monday
schedule.every().wednesday.at("13:15").do(job)  # Send email every Wednesday at 13:15
schedule.every().day.at("12:42").do(job)  # Send email every day at 12:42 PM
schedule.every(5).minutes.at(":17").do(job)  # Send email every 5 minutes at 17 seconds

# Run the schedule in a loop
while True:
    schedule.run_pending()
    time.sleep(1)
