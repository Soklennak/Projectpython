import smtplib
from email.mime.text import MIMEText

# Set up the server
smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
smtpObj.starttls()

# Use your actual Gmail address and either your regular password (if less secure app access is enabled)
# or the app password generated if you have 2FA enabled
smtpObj.login('soklen.nak@student.passerellesnumeriques.org', 'pgfp gxpz fjbr qkll')

# Construct the email
msg = MIMEText('This is the body of the email testing sending')
msg['Subject'] = 'Test Email sreyvang '
msg['From'] = 'soklen.nak@student.passerellesnumeriques.org'
msg['To'] = 'soklennak01@gmail.com'

# Send the email
smtpObj.sendmail('soklen.nak@student.passerellesnumeriques.org', ['soklennak01@gmail.com'], msg.as_string())
smtpObj.quit()
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Email content
sender = 'soklen.nak@student.passerellesnumeriques.org'
receivers = ['soklennak01@gmail.com']

# Create message container - the correct MIME type is multipart/alternative.
msg = MIMEMultipart('alternative')
msg['From'] = 'From Person <soklrn.nak@student.passerellesnumeriques.org>'
msg['To'] = 'To Person <soklennak01@gmail.com>'
msg['Subject'] = 'SMTP HTML e-mail test'

# HTML message content
html = """\
<html>
  <head></head>
  <body>
    <p>This is an e-mail message to be sent in <b>HTML format</b></p>
    <p><b>This is HTML message.</b></p>
    <h1>This is headline.</h1>
  </body>
</html>
"""
# Attach HTML content to the email
part2 = MIMEText(html, 'html')
msg.attach(part2)

# Connect to SMTP server and send email
try:
   smtpObj = smtplib.SMTP('localhost')
   smtpObj.sendmail(sender, receivers, msg.as_string())
   print("Successfully sent email")
except smtplib.SMTPException as e:
   print(f"Error: unable to send email. Error message: {str(e)}")


