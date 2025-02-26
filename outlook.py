import csv
import win32com.client

# Read email addresses from CSV
email_addresses = []
with open("contacts.csv", "r") as file:
    csv_reader = csv.DictReader(file)
    for row in csv_reader:
        email_addresses.append(row["email"])  # Assuming 'email' is your column header

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Your email template
subject = "Important Announcement"
body = """
Hello,

This is our form letter with important information.

Best regards,
Your Name
"""

# Send emails
for email in email_addresses:
    mail = outlook.CreateItem(0)  # 0 represents an email item
    mail.Subject = subject
    mail.Body = body
    mail.To = email

    # Uncomment to actually send the email
    # mail.Send()

    # Or display it for review before sending
    mail.Display()

print(f"Process completed for {len(email_addresses)} recipients")
