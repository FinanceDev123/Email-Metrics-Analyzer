import imaplib
import email
from email.header import decode_header
import datetime
import os
import getpass
import re
from collections import Counter
from datetime import timedelta
from email import message_from_bytes

# Function to connect to the email server
def connect_to_email(username, password, imap_server="imap.gmail.com"):
    print("Attempting to connect to the email server...")  # Move it outside the try block
    
    try:
        mail = imaplib.IMAP4_SSL(imap_server)
        print("Successfully created SSL connection.")
        
        mail.login(username, password)
        print("Login successful.")
        
        # Check the connection status using NOOP
        status, response = mail.noop()
        print(f"NOOP response: {response}")  # Print the response from the NOOP command
        
        if status == 'OK':
            print("Successfully connected to the email server.\n\n")
        else:
            print(f"Connection check failed: {response}")
        
        return mail
    except Exception as e:
        print(f"Failed to connect to email: {e}")
        return None

# Function to decode HTML body content into plain text
def get_plain_text_body(msg):
    # Check if the email is multipart
    if msg.is_multipart():
        # Iterate over each part
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))

            if "attachment" not in content_disposition:
                if content_type == "text/plain":  # Plain text part
                    return part.get_payload(decode=True).decode()  # Return the plain text part
                elif content_type == "text/html":  # HTML part
                    return part.get_payload(decode=True).decode()  # Return the HTML part (for now, you can add parsing logic later)
    else:
        # If it's not multipart, just return the payload
        return msg.get_payload(decode=True).decode()

# Function to search emails by date and keywords
def search_emails(mail, folder="inbox", date_from=None, date_to=None, keywords=None):
    try:
        mail.select(folder)  # Ensure the folder is correctly selected

        # Format the date range for the search
        next_day = date_to + timedelta(days=1)
        search_query = []
        if date_from:
            search_query.append(f"SINCE {date_from.strftime('%d-%b-%Y')}")
        if date_to:
            search_query.append(f"BEFORE {next_day.strftime('%d-%b-%Y')}")

        # Execute the search query
        status, messages = mail.search(None, *search_query)
        if status != "OK":
            print("No messages found.")
            return []

        matching_emails = []
        for num in messages[0].split():
            status, msg_data = mail.fetch(num, '(RFC822)')
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    subject, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(encoding if encoding else "utf-8")
                    body = get_plain_text_body(msg)  # Get plain text body

                    if keywords:
                        if any(keyword.lower() in (subject or "").lower() for keyword in keywords):
                            matching_emails.append((subject, msg.get("From"), msg.get("Date"), body))
                    else:
                        matching_emails.append((subject, msg.get("From"), msg.get("Date"), body))
        return matching_emails
    except Exception as e:
        print(f"Error searching emails: {e}")
        return []

if __name__ == "__main__":
    EMAIL = input("Enter your email: ")
    PASSWORD = getpass.getpass("Enter your password: ")

    DATE_FROM = datetime.date(2025, 1, 1)  # Change to the desired starting date of the search - Format: Year/Month/Day
    DATE_TO = datetime.date(2025, 1, 24)  # Input today's date
    KEYWORDS = ["Keyword 1", "Keyword 2"] # Input the desired Keywords

    mail = connect_to_email(EMAIL, PASSWORD)
    if mail:
        # Search emails
        emails = search_emails(mail, folder="inbox", date_from=DATE_FROM, date_to=DATE_TO, keywords=KEYWORDS)

        # Print matching emails
        if emails:
            for subject, sender, date, body in emails:  # Include 'body' in the tuple unpacking
                print ("\033[4mResults Received:\033[0m\n")
                print(f"Subject: {subject}\nFrom: {sender}\nDate: {date}\nBody: {body[:100]}...")  # Print first 100 characters of body
        else:
            print("No matching emails found.")

        # Metrics
        print("\nMetrics:")
        print(f"Total emails found: {len(emails)}")
        sender_counts = Counter(sender for _, sender, _, _ in emails)  # Adjusted for 4-tuple
        print("\nEmails by sender:")
        for sender, count in sender_counts.items():
            print(f"{sender}: {count} emails")
        date_counts = Counter(date for _, _, date, _ in emails)  # Adjusted for 4-tuple
        print("\nEmails by date:")
        for date, count in date_counts.items():
            print(f"{date}: {count} emails")
        
        # Keyword frequency analysis (check subject and body)
        keyword_counts = Counter()
        for subject, _, _, body in emails:  # Adjusted for 'body'
            # Check in subject
            for keyword in KEYWORDS:
                if keyword.lower() in (subject or "").lower():
                    keyword_counts[keyword] += 1
            # Check in body
            for keyword in KEYWORDS:
                if keyword.lower() in (body or "").lower():
                    keyword_counts[keyword] += 1
        
        print("\nKeyword frequency:")
        for keyword, count in keyword_counts.items():
            print(f"{keyword}: {count} occurrences")
        
        print("\nSummary:")
        print(f"Total emails found: {len(emails)}")
        if sender_counts:
            top_sender = max(sender_counts, key=sender_counts.get)
            print(f"Top sender: {top_sender} ({sender_counts[top_sender]} emails)")
        if date_counts:
            top_date = max(date_counts, key=date_counts.get)
            print(f"Most active date: {top_date} ({date_counts[top_date]} emails)")
        if keyword_counts:
            top_keyword = max(keyword_counts, key=keyword_counts.get)
            print(f"Most mentioned keyword: {top_keyword} ({keyword_counts[top_keyword]} occurrences)")
            

# Logout
try:
    mail.logout()
    print("Successfully logged out from the email server.")
except Exception as e:
    print(f"Error logging out: {e}")

# Reset inputs
EMAIL = ""
PASSWORD = ""
DATE_FROM = None
DATE_TO = None
KEYWORDS = []

# Print to confirm reset
print("Inputs have been successfully reset.")

# Print matching emails
if emails:
    for subject, sender, date, body in emails:  # Now unpack 'body' as well
        print ("\033[4mTo conclude, here is a summary of the results:\033[0m\n")
    
        print(f"Subject: {subject}\nFrom: {sender}\nDate: {date}\nBody: {body[:100]}...")  # Print first 100 characters of body
else:
    print("No matching emails found.")

# Metrics and other output...

# End of the script
print("\n\n\033[1m\033[32mThank you for using the email search tool!\033[0m")
print("\033[1m\033[32mHave a wonderful day!\033[0m")

# Initial upload of email search tool code
