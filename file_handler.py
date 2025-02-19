import os
from typing import Set

class FileHandler:
    def __init__(self, processed_emails_file: str = "processed_emails.txt"):
        self.processed_emails_file = processed_emails_file

    def load_processed_emails(self) -> Set[str]:
        if os.path.exists(self.processed_emails_file):
            with open(self.processed_emails_file, "r") as f:
                return set(line.strip() for line in f)
        return set()

    def save_processed_email(self, email_id: str):
        with open(self.processed_emails_file, "a") as f:
            f.write(email_id + "\n") 