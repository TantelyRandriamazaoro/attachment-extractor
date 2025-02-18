import msal
import requests
from dotenv import load_dotenv
import os
import base64
from datetime import datetime
import pytz

# Load environment variables from .env file
load_dotenv()

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
SENDERS = os.getenv("SENDERS").split(",")  # Comma-separated list of email addresses
MAILBOX = os.getenv("MAILBOX")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]

# Initialize the MSAL confidential client application
app = msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)

# Acquire token for client
token_result = app.acquire_token_for_client(scopes=SCOPES)

if "access_token" in token_result:
    access_token = token_result["access_token"]
    print("Access Token Acquired")

    # Set the headers for requests
    headers = {"Authorization": f"Bearer {access_token}"}

    # Get current local time in the desired timezone (e.g., GMT+3)
    local_timezone = pytz.timezone('Etc/GMT-3')
    now = datetime.now(local_timezone)

    # Set the time to midnight (start of the day) in the current timezone
    start_of_day_local = now.replace(hour=0, minute=0, second=0, microsecond=0)

    # Convert the start of the day to UTC
    start_of_day_utc = start_of_day_local.astimezone(pytz.utc)

    # Format the time in the required format
    formatted_time = start_of_day_utc.strftime('%Y-%m-%dT%H:%M:%S') + 'Z'

    print(f"Retrieving emails received today ({formatted_time})")

    # Create a folder with the current date as its name under /attachments
    folder_name = now.strftime("%Y-%m-%d")
    os.makedirs(f"attachments/{folder_name}", exist_ok=True)

    # Filter emails received today
    filter_query = f"receivedDateTime ge {formatted_time}"

    # Retrieve emails
    response = requests.get(f"https://graph.microsoft.com/v1.0/users/{MAILBOX}/messages/?$top=50&$filter={filter_query}", headers=headers)
    emails = response.json()

    if "value" in emails:
        for email in emails["value"]:
            sender_email = email.get("from", {}).get("emailAddress", {}).get("address")
            subject = email.get("subject", "")

            print('-'*50)
            
            # Check if the email sender is one of the specified senders
            if sender_email in SENDERS:
                email_id = email["id"]
                
                print(f"Processing email from {sender_email} ({subject})")

                # Fetch attachments for the email
                attachment_url = f"https://graph.microsoft.com/v1.0/users/{MAILBOX}/messages/{email_id}/attachments"
                attachment_response = requests.get(attachment_url, headers=headers)
                attachments = attachment_response.json()

                if "value" in attachments:
                    for attachment in attachments["value"]:
                        attachment_id = attachment["id"]
                        attachment_name = attachment["name"]
                        attachment_content = attachment.get("contentBytes", "")

                        # Save the attachment to a file in the created folder
                        if attachment_content:
                            decoded_content = base64.b64decode(attachment_content)
                            file_path = os.path.join('attachments/' + folder_name, attachment_name)

                            with open(file_path, "wb") as f:
                                f.write(decoded_content)
                            print(f"Attachment {attachment_name} saved to {file_path}.")
                else:
                    print("No attachments found for this email.")
            else:
                print(f"Email from {sender_email} skipped.")
    else:
        print("No emails found.")
else:
    print("Error:", token_result.get("error_description"))
