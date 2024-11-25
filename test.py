import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os


<<<<<<< HEAD
def schedule_email(date_time):
    while True:
        # Check if current time matches the specified time
        now = datetime.now()
        if now >= date_time:
            
            # Email Sending Function
            def send_email(to_email, subject, body):
                try:
                    # Create Email Message
                    message = MIMEMultipart()
                    message["From"] = email_user
                    message["To"] = to_email
                    message["Subject"] = subject
                    message.attach(MIMEText(body, "plain"))
=======
def send_bulk_emails(smtp_server, port, sender_email, sender_password, subject, body, excel_file, attachment_path):
    try:
        # Read the Excel file
        df = pd.read_excel(excel_file)
        if 'Email' not in df.columns:
            messagebox.showerror("Error", "The Excel file must contain a column named 'Email'.")
            return
        
        for index, row in df.iterrows():
            recipient = row['Email']
            send_email(smtp_server, port, sender_email, sender_password, recipient, subject, body, attachment_path)
        
        messagebox.showinfo("Success", "Emails sent successfully!")
    
    except Exception as e:
        messagebox.showerror("Error", f"Failed to send emails: {e}")

def send_email(smtp_server, port, sender_email, sender_password, recipient, subject, body, attachment_path):
    try:
        # Create the email object
        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = recipient
        message['Subject'] = subject
        
        # Add the HTML body
        message.attach(MIMEText(body, 'html'))
        
        # Add an attachment if provided
        if attachment_path and os.path.exists(attachment_path):
            with open(attachment_path, 'rb') as file:
                mime_base = MIMEBase('application', 'octet-stream')
                mime_base.set_payload(file.read())
                encoders.encode_base64(mime_base)
                mime_base.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_path)}')
                message.attach(mime_base)
        
        # Send the email
        with smtplib.SMTP(smtp_server, port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(message)
    
    except Exception as e:
        print(f"Failed to send email to {recipient}: {e}")

def browse_file(entry):
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    entry.delete(0, tk.END)
    entry.insert(0, filename)
>>>>>>> 4261a733abee4efc11359d06e0ca93e61d6269c5

def browse_attachment(entry):
    filename = filedialog.askopenfilename()
    entry.delete(0, tk.END)
    entry.insert(0, filename)

def send_emails():
    smtp_server = smtp_entry.get()
    port = int(port_entry.get())
    sender_email = sender_entry.get()
    sender_password = password_entry.get()
    subject = subject_entry.get()
    body = body_text.get("1.0", tk.END)
    excel_file = file_entry.get()
    attachment_path = attachment_entry.get()
    
    if not all([smtp_server, port, sender_email, sender_password, subject, body, excel_file]):
        messagebox.showerror("Error", "Please fill in all required fields.")
        return
    
    send_bulk_emails(smtp_server, port, sender_email, sender_password, subject, body, excel_file, attachment_path)

<<<<<<< HEAD
# Set the date and time for sending the email
# Example: November 19, 2024, at 14:9
scheduled_time = datetime(2024, 11, 24, 3, 34)

# Start the scheduler
print(f"Email will be send at {scheduled_time}")
schedule_email(scheduled_time)
=======
# Create the GUI
root = tk.Tk()
root.title("Email Automation System")
root.geometry("500x600")

# SMTP Server
tk.Label(root, text="SMTP Server:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
smtp_entry = tk.Entry(root, width=40)
smtp_entry.grid(row=0, column=1, padx=10, pady=5)
>>>>>>> 4261a733abee4efc11359d06e0ca93e61d6269c5

# Port
tk.Label(root, text="Port:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
port_entry = tk.Entry(root, width=40)
port_entry.grid(row=1, column=1, padx=10, pady=5)

# Sender Email
tk.Label(root, text="Sender Email:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
sender_entry = tk.Entry(root, width=40)
sender_entry.grid(row=2, column=1, padx=10, pady=5)

# Sender Password
tk.Label(root, text="Password:").grid(row=3, column=0, sticky="w", padx=10, pady=5)
password_entry = tk.Entry(root, width=40, show="*")
password_entry.grid(row=3, column=1, padx=10, pady=5)

# Subject
tk.Label(root, text="Subject:").grid(row=4, column=0, sticky="w", padx=10, pady=5)
subject_entry = tk.Entry(root, width=40)
subject_entry.grid(row=4, column=1, padx=10, pady=5)

# Body
tk.Label(root, text="Body:").grid(row=5, column=0, sticky="nw", padx=10, pady=5)
body_text = tk.Text(root, width=40, height=10)
body_text.grid(row=5, column=1, padx=10, pady=5)

# Excel File
tk.Label(root, text="Excel File:").grid(row=6, column=0, sticky="w", padx=10, pady=5)
file_entry = tk.Entry(root, width=30)
file_entry.grid(row=6, column=1, padx=10, pady=5, sticky="w")
file_button = tk.Button(root, text="Browse", command=lambda: browse_file(file_entry))
file_button.grid(row=6, column=1, padx=10, pady=5, sticky="e")

# Attachment
tk.Label(root, text="Attachment:").grid(row=7, column=0, sticky="w", padx=10, pady=5)
attachment_entry = tk.Entry(root, width=30)
attachment_entry.grid(row=7, column=1, padx=10, pady=5, sticky="w")
attachment_button = tk.Button(root, text="Browse", command=lambda: browse_attachment(attachment_entry))
attachment_button.grid(row=7, column=1, padx=10, pady=5, sticky="e")

# Send Button
send_button = tk.Button(root, text="Send Emails", command=send_emails)
send_button.grid(row=8, column=0, columnspan=2, pady=20)

root.mainloop()
