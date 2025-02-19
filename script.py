from dotenv import load_dotenv
import os
from auth import MSGraphAuth
from email_processor import EmailProcessor
from file_handler import FileHandler

def main():
    # Load environment variables
    load_dotenv()
    
    # Initialize components
    auth = MSGraphAuth()
    file_handler = FileHandler()
    
    try:
        # Get access token
        access_token = auth.get_access_token()
        print("Access Token Acquired")

        # Initialize email processor
        mailbox = os.getenv("MAILBOX")
        email_processor = EmailProcessor(access_token, mailbox)

        # Load processed emails and process new ones
        processed_emails = file_handler.load_processed_emails()
        email_processor.process_emails(processed_emails)

    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()
