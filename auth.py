import os
import msal

class MSGraphAuth:
    def __init__(self):
        self.client_id = os.getenv("CLIENT_ID")
        self.client_secret = os.getenv("CLIENT_SECRET")
        self.tenant_id = os.getenv("TENANT_ID")
        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        self.scopes = ["https://graph.microsoft.com/.default"]
        
        self.app = msal.ConfidentialClientApplication(
            self.client_id, 
            authority=self.authority, 
            client_credential=self.client_secret
        )

    def get_access_token(self):
        token_result = self.app.acquire_token_for_client(scopes=self.scopes)
        if "access_token" in token_result:
            return token_result["access_token"]
        raise Exception(f"Error getting token: {token_result.get('error_description')}") 