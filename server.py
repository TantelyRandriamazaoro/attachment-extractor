from flask import Flask, redirect, request, session, url_for
import msal
import os
from dotenv import load_dotenv
import requests

app = Flask(__name__)
app.secret_key = os.urandom(24)  # Required for session management
load_dotenv()

# OAuth configuration
CLIENT_ID = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
AUTHORITY = "https://login.microsoftonline.com/common"
REDIRECT_URI = "http://localhost:3000/token"
SCOPE = ["User.Read", "Mail.Read"]  # Add any other required scopes

@app.route('/')
def login():
    # Initialize MSAL client
    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )
    
    # Get auth URL
    auth_url = msal_app.get_authorization_request_url(
        SCOPE,
        redirect_uri=REDIRECT_URI,
        state=None
    )
    
    return redirect(auth_url)

@app.route('/token')
def get_token():
    if "error" in request.args:
        return f"Error: {request.args['error']}"
    
    if "code" not in request.args:
        return "No code received."
    
    # Initialize MSAL client
    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )
    
    # Get token using the authorization code
    result = msal_app.acquire_token_by_authorization_code(
        request.args['code'],
        scopes=SCOPE,
        redirect_uri=REDIRECT_URI
    )
    
    if "error" in result:
        return f"Error: {result['error']}"
    
    # Store the token in session
    session['access_token'] = result['access_token']
    
    return redirect(url_for('extract'))

@app.route('/extract')
def extract():
    if 'access_token' not in session:
        return redirect(url_for('login'))
    
    access_token = session['access_token']
    
    try:
        # Initialize components
        file_handler = FileHandler()
        
        # Initialize email processor with the access token from OAuth
        mailbox = os.getenv("MAILBOX")
        email_processor = EmailProcessor(access_token, mailbox)

        # Load processed emails and process new ones
        processed_emails = file_handler.load_processed_emails()
        email_processor.process_emails(processed_emails)
        
        return "Extraction completed successfully!"
        
    except Exception as e:
        return f"Error during extraction: {str(e)}"

if __name__ == '__main__':
    app.run(host='localhost', port=3000, debug=True) 