import os
import requests
import base64
from typing import Set

class EmailProcessor:
    def __init__(self, access_token: str, mailbox: str):
        self.headers = {"Authorization": f"Bearer {access_token}"}
        self.mailbox = mailbox
        self.base_url = "https://graph.microsoft.com/v1.0"

    def process_emails(self, processed_emails: Set[str]):
        link = f"{self.base_url}/users/{self.mailbox}/messages/?$top=50"
        
        while link:
            try:
                emails = self._get_emails(link)
                if not emails.get("value"):
                    print("No emails found.")
                    break

                print(f"Retrieved {len(emails['value'])} emails. Processing...")
                
                for email in emails["value"]:
                    self._process_single_email(email, processed_emails)

                link = emails.get("@odata.nextLink")
                if link:
                    print("Processing next page...")

            except requests.exceptions.RequestException as e:
                print(f"Error during request: {e}")
                break

    def _get_emails(self, url: str) -> dict:
        response = requests.get(url, headers=self.headers)
        response.raise_for_status()
        return response.json()

    def _process_single_email(self, email: dict, processed_emails: Set[str]):
        email_id = email["id"]
        sender_email = email.get("from", {}).get("emailAddress", {}).get("address")
        subject = email.get("subject", "")

        print("-" * 50)
        if email_id in processed_emails:
            print(f"Email '{subject}' from {sender_email} already processed. Skipping...")
            return

        print(f"Processing email from \033[34m{sender_email}\033[0m with subject \033[34m{subject}\033[0m")
        
        attachments = self._get_attachments(email_id)
        if attachments.get("value"):
            self._save_attachments(attachments["value"], sender_email)
        else:
            print("No attachments found for this email.")

        return email_id

    def _get_attachments(self, email_id: str) -> dict:
        url = f"{self.base_url}/users/{self.mailbox}/messages/{email_id}/attachments"
        response = requests.get(url, headers=self.headers)
        response.raise_for_status()
        return response.json()

    def _save_attachments(self, attachments: list, sender_email: str):
        os.makedirs(f"attachments/{sender_email}", exist_ok=True)

        for attachment in attachments:
            attachment_name = attachment["name"]
            file_path = os.path.join('attachments', sender_email, attachment_name)

            if not attachment_name.startswith("Outlook-") and not os.path.exists(file_path):
                attachment_content = attachment.get("contentBytes", "")
                if attachment_content:
                    self._save_attachment_file(file_path, attachment_content)
                    print(f"Attachment {attachment_name} saved to \033[33m{file_path}\033[0m")
            else:
                if os.path.exists(file_path):
                    print(f"Attachment {attachment_name} already exists in {sender_email} folder. Skipping...")

    def _save_attachment_file(self, file_path: str, content: str):
        decoded_content = base64.b64decode(content)
        with open(file_path, "wb") as f:
            f.write(decoded_content) 