import imaplib
import email
from email.header import decode_header
import os
import sys
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import ssl
import re

# Ensure a directory path is provided for saving attachments
if len(sys.argv) != 2:
    print("Usage: AttachmentDownloader.py <path_to_save_attachments>")
    sys.exit(1)
attachment_dir = sys.argv[1]


# IMAP server configuration (placeholders to be replaced with actual values)
imap_host = 'IMAP Server'
imap_user = 'IMAP Username'
imap_pass = 'IMAP Password'
imap_from = 'From Name'

# List of file extensions not allowed to be saved
DISALLOWED_EXTENSIONS = [
    '.exe', '.bat', '.cmd', '.scr', '.ps1', '.psm1', '.psd1',
    '.vbs', '.vbe', '.js', '.jse', '.wsf', '.wsh', '.msc',
    '.com', '.reg', '.dll', '.lnk', '.zip', '.rar', '.7z',
    '.tar', '.gz', '.jar'
]

def is_allowed_attachment(filename):
    """Determines if a file's extension is allowed."""
    extension = os.path.splitext(filename)[1].lower()
    return extension not in DISALLOWED_EXTENSIONS

# Adjust SSL settings for IMAP and SMTP
ssl_context = ssl.create_default_context()
ssl_context.set_ciphers('DEFAULT@SECLEVEL=1')
ssl_context.check_hostname = False
ssl_context.verify_mode = ssl.CERT_NONE

# Connect to IMAP server
mail = imaplib.IMAP4_SSL(imap_host, ssl_context=ssl_context)
mail.login(imap_user, imap_pass)
mail.select('inbox')

def send_confirmation(sender_email, subject, attachments):
    """Sends a confirmation email to the sender detailing the saved attachments using the same IMAP server settings."""
    from_name = imap_from

    # Construct the email message
    msg = MIMEMultipart()
    msg['From'] = f"{from_name} <{imap_user}>"
    msg['To'] = sender_email
    msg['Subject'] = f"Dropbox Confirmation: {subject}"
    body = (f"Your email with the subject '{subject}' was processed. Attachments saved:\n"
            + '\n'.join(attachments) +
            "\n\nNote: Emails with disallowed extensions are not saved.")
    msg.attach(MIMEText(body, 'plain'))

    # Send the email using the same IMAP server for SMTP
    server = smtplib.SMTP_SSL(imap_host, context=ssl_context)
    server.login(imap_user, imap_pass)
    server.sendmail(imap_user, sender_email, msg.as_string())
    server.quit()

def sanitize_filename(filename):
    """Cleans filenames of unwanted characters."""
    sanitized = re.sub(r'[<>:"/\\|?*\r\n]+', '_', filename)
    sanitized = sanitized.strip().strip('.')
    return sanitized

def get_available_filename(directory, filename):
    """Generates a unique filename if there is a name clash."""
    base, extension = os.path.splitext(filename)
    counter = 1
    new_filename = filename
    while os.path.exists(os.path.join(directory, new_filename)):
        new_filename = f"{base}_{str(counter).zfill(3)}{extension}"
        counter += 1
    return new_filename

# Process unread emails
status, messages = mail.search(None, '(UNSEEN)')
if status != 'OK':
    print("No unread emails found.")
    mail.logout()
    sys.exit()

messages = messages[0].split()
for mail_id in messages:
    res, msg_data = mail.fetch(mail_id, '(RFC822)')
    if res != 'OK':
        continue
    msg = email.message_from_bytes(msg_data[0][1])
    subject = decode_header(msg.get("Subject"))[0][0]
    if isinstance(subject, bytes):
        subject = subject.decode()
    attachments_saved = []

    for part in msg.walk():
        if part.get_content_maintype() == 'multipart' or not part.get('Content-Disposition'):
            continue
        filename = part.get_filename()
        if filename and is_allowed_attachment(filename):
            filename = decode_header(filename)[0][0]
            if isinstance(filename, bytes):
                filename = filename.decode()
            filename = sanitize_filename(filename)
            filename = get_available_filename(attachment_dir, filename)
            filepath = os.path.join(attachment_dir, filename)
            if not os.path.exists(attachment_dir):
                os.makedirs(attachment_dir)
            with open(filepath, "wb") as f:
                f.write(part.get_payload(decode=True))
            attachments_saved.append(filename)
    if attachments_saved:
        send_confirmation(msg['From'], subject, attachments_saved)

    mail.store(mail_id, '+FLAGS', '\\Seen')

mail.close()
mail.logout()
print("Finished processing emails.")
