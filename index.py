import msal
import requests
import os
from flask import Flask, redirect, request, session, url_for, send_file
from dotenv import load_dotenv
import base64
from datetime import datetime
import pytz
import math
import zipfile
from io import BytesIO

load_dotenv()

# Set up your app registration details
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
MAILBOX = os.getenv("MAILBOX")
SENDERS = os.getenv("SENDERS").split(",")  # Comma-separated list of email addresses
AUTHORITY = f'https://login.microsoftonline.com/{TENANT_ID}'
REDIRECT_URI = 'http://localhost:3000/token'

# Scopes (what permissions you're requesting)
SCOPES = ['User.Read', 'Mail.Read']

# Initialize the Flask app
app = Flask(__name__)
app.secret_key = os.urandom(24)

# Initialize MSAL app
msal_app = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET
)

# Redirect to Microsoft's OAuth2 Authorization Endpoint
@app.route('/')
def index():
    auth_url = msal_app.get_authorization_request_url(SCOPES, redirect_uri=REDIRECT_URI)
    print(auth_url)
    return redirect(auth_url)

# Handle the redirect from the Microsoft OAuth2 Authorization Server
@app.route('/token')
def get_token():
    code = request.args.get('code')
    
    if not code:
        return "Authorization failed", 400
    
    # Exchange the authorization code for an access token
    result = msal_app.acquire_token_by_authorization_code(
        code,
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI
    )

    if "access_token" in result:
        # Save the access token in the session for use in other parts of your app
        session['access_token'] = result['access_token']
        return redirect(url_for('download'))
    else:
        return f"Failed to acquire token: {result.get('error_description')}", 400

@app.route('/download')
def download():
    # Set the headers for requests
    headers = {"Authorization": f"Bearer {session['access_token']}"}

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
        processed_attachments = []

        for email in emails["value"]:
            sender_email = email.get("from", {}).get("emailAddress", {}).get("address")
            subject = email.get("subject", "")
            
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
                        attachment_name = attachment["name"]
                        attachment_content = attachment.get("contentBytes", "")

                        # Save the attachment to a file in the created folder
                        if attachment_content:
                            decoded_content = base64.b64decode(attachment_content)
                            file_path = os.path.join('attachments', folder_name, attachment_name)

                            with open(file_path, "wb") as f:
                                f.write(decoded_content)
                            print(f"Attachment {attachment_name} saved to {file_path}.")

                        processed_attachments.append({
                            "name": attachment_name,
                            "path": file_path
                        })
                else:
                    print("No attachments found for this email.")
            else:
                print(f"Email from {sender_email} skipped.")

        if processed_attachments:
            # Create a ZIP file in memory
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                # Add each file in the processed_attachments list to the ZIP file
                for attachment in processed_attachments:
                    attachment_path = attachment['path']
                    zip_file.write(attachment_path, arcname=attachment['name'])
            
            # Move the buffer position to the start
            zip_buffer.seek(0)

            # Send the ZIP file as a download
            return send_file(zip_buffer,
                             as_attachment=True,
                             download_name=f"attachments_{folder_name}.zip",
                             mimetype="application/zip")
        else:
            return "No attachments found to download."

    else:
        return "No emails found."


# Run the Flask app
if __name__ == '__main__':
    app.run(debug=True, port=3000)

