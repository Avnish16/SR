import smtplib
import random
import string
import time
from email.message import EmailMessage
from pathlib import Path

# ==============================================================
# CONFIGURATION
# ==============================================================

SENDER_EMAIL = "xyz@gmail.com"
APP_PASSWORD = "pass"  # create in Outlook account security
RECEIVER_EMAIL = "abc@outlook.com"
ATTACHMENT_PATH = "OriginalMsg.eml"  # must exist in same folder or specify full path
TOTAL_EMAILS = 5

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# ==============================================================

def generate_interaction_id():
    """Generate a unique interaction ID (same pattern as your example)."""
    prefix = "G502R511MG805KEK"
    random_part = ''.join(random.choices(string.ascii_uppercase + string.digits, k=14))
    return f"{prefix}.{random_part}."

def generate_parent_interaction_id():
    """Generate a unique parent interaction ID (like your pattern)."""
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=14))

def build_email(interaction_id, parent_id):
    """Constructs one email with the specified IDs."""
    msg = EmailMessage()
    msg["From"] = SENDER_EMAIL
    msg["To"] = RECEIVER_EMAIL
    msg["Subject"] = f"Interaction {interaction_id}"

    # Email body with dynamic IDs
    body = f"""Forward To mailbox 
---- Customer Email attached as file OriginalMsg.eml ----

{interaction_id}
------------------------------------------------------

Hier ist Platz für Ihre Rückmeldung.

------------------------------------------------------
{interaction_id}


<contact-form>False</contact-form>
<parent-interaction-id>{parent_id}</parent-interaction-id>
<customer-email-address>xyz@outlook.com</customer-email-address>
<category>Allgemein</category>
<subcategory>None</subcategory>

___________________________________________________
Mit freundlichen Grüßen
E-Mail-Team

E-Mail-Service
Bank | India
ABC AG
abc@outlook.com
"""

    msg.set_content(body)

    # Attach .eml file
    with open(ATTACHMENT_PATH, "rb") as f:
        msg.add_attachment(f.read(), maintype="message", subtype="rfc822", filename="OriginalMsg.eml")

    return msg

# ==============================================================

def send_bulk_emails():
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.login(SENDER_EMAIL, APP_PASSWORD)
        print("Logged in to Gmail SMTP")

        # --- TIMER START ---
        start_time = time.time()
        print("Starting email sending...\n")

        for i in range(TOTAL_EMAILS):
            interaction_id = generate_interaction_id()
            parent_id = generate_parent_interaction_id()
            email = build_email(interaction_id, parent_id)

            smtp.send_message(email)
            print(f"[{i+1}/{TOTAL_EMAILS}] Sent: {interaction_id}")

        # --- TIMER END ---
        end_time = time.time()
        elapsed = end_time - start_time
        print("\nAll emails sent successfully!")
        print(f"Total time taken: {elapsed:.2f} seconds ({elapsed/60:.2f} minutes)")

# ==============================================================

if __name__ == "__main__":
    send_bulk_emails()
