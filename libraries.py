import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Email details
sender_email = "soklennak01@gmail.com"
receiver_email = "soklen.nak@student.passerellesnumeriques.org"
password = "#lenylove*918"  # Use App Password if 2FA is enabled

# Create a multipart message
msg = MIMEMultipart("alternative")
msg['From'] = sender_email
msg['To'] = receiver_email
msg['Subject'] = "Test Email with Plain Text and HTML"

# Create the plain-text and HTML versions of the message
text = "Hello, this is the plain text version of the email."
html = """\
<html>
  <body>
    <p>Hello,<br>
       This is the <b>HTML</b> version of the email!</p>
  </body>
</html>
"""

# Attach both the plain text and HTML parts to the email
msg.attach(MIMEText(text, 'plain'))
msg.attach(MIMEText(html, 'html'))

# Set up the server and send email
try:
    # Connect to Gmail's SMTP server (for Gmail use 'smtp.gmail.com' and port 587)
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()  # Secure connection
    server.login(sender_email, password)  # Log in to the server
    
    # Send the email
    text = msg.as_string()
    server.sendmail(sender_email, receiver_email, text)
    
    print("Email sent successfully!")
    
except Exception as e:
    print(f"Failed to send email: {e}")
    
finally:
    server.quit()  # Close the connection to the server
