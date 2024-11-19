import smtplib
import tkinter as tk
from tkinter import messagebox, filedialog
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

def send_email():
    # Get user inputs
    sender_email = entry_sender.get()
    receiver_email = entry_receiver.get()
    subject = entry_subject.get()
    body = text_message.get("1.0", tk.END)

    # Email credentials and SMTP server details
    password = entry_password.get()
    smtp_server = "smtp.gmail.com"
    smtp_port = 587

    # Create the email message container (MIME multipart)
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = receiver_email
    message['Subject'] = subject
    message.attach(MIMEText(body, 'plain'))

    # Add file attachment if any
    if attachment_path:
        try:
            attachment = open(attachment_path, "rb")
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename={attachment_path.split('/')[-1]}")
            message.attach(part)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to attach file: {str(e)}")
            return

    try:
        # Set up the SMTP server
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # Secure connection
        server.login(sender_email, password)

        # Send the email
        server.sendmail(sender_email, receiver_email, message.as_string())
        server.quit()

        # Notify user that email is sent successfully
        messagebox.showinfo("Success", "Email sent successfully!")

    except Exception as e:
        messagebox.showerror("Error", f"Failed to send email: {str(e)}")

def browse_file():
    """Open file dialog to select a file to attach."""
    global attachment_path
    attachment_path = filedialog.askopenfilename(title="Select a file to attach", 
                                                filetypes=[("All files", "*.*")])
    if attachment_path:
        label_attachment.config(text=f"Attachment: {attachment_path.split('/')[-1]}")
    else:
        label_attachment.config(text="No file attached")

# Set up the GUI window
root = tk.Tk()
root.title("Email Sender with Attachment")

# Sender email label and entry
label_sender = tk.Label(root, text="Sender Email:")
label_sender.grid(row=0, column=0, padx=10, pady=5)
entry_sender = tk.Entry(root, width=30)
entry_sender.grid(row=0, column=1, padx=10, pady=5)

# Sender email password entry (hidden)
label_password = tk.Label(root, text="Sender Password:")
label_password.grid(row=1, column=0, padx=10, pady=5)
entry_password = tk.Entry(root, width=30, show="*")
entry_password.grid(row=1, column=1, padx=10, pady=5)

# Receiver email label and entry
label_receiver = tk.Label(root, text="Receiver Email:")
label_receiver.grid(row=2, column=0, padx=10, pady=5)
entry_receiver = tk.Entry(root, width=30)
entry_receiver.grid(row=2, column=1, padx=10, pady=5)

# Subject label and entry
label_subject = tk.Label(root, text="Subject:")
label_subject.grid(row=3, column=0, padx=10, pady=5)
entry_subject = tk.Entry(root, width=30)
entry_subject.grid(row=3, column=1, padx=10, pady=5)

# Message Textbox
label_message = tk.Label(root, text="Message:")
label_message.grid(row=4, column=0, padx=10, pady=5)
text_message = tk.Text(root, height=10, width=30)
text_message.grid(row=4, column=1, padx=10, pady=5)

# File Attachment Button
attachment_path = ""  # To store the file path
button_browse = tk.Button(root, text="Attach File", command=browse_file)
button_browse.grid(row=5, column=0, padx=10, pady=5)

# Label to display selected attachment file name
label_attachment = tk.Label(root, text="No file attached")
label_attachment.grid(row=5, column=1, padx=10, pady=5)

# Send Button
send_button = tk.Button(root, text="Send Email", command=send_email)
send_button.grid(row=6, column=1, pady=10)

# Start the GUI event loop
root.mainloop()
