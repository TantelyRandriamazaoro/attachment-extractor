
# Microsoft Graph Email Attachment Downloader

This Flask web application allows you to authenticate using Microsoft OAuth2, retrieve emails from a specified mailbox, and download email attachments from specific senders. The attachments are bundled into a ZIP file for easy download.

## Prerequisites

To run this script, you need the following:

- A Microsoft Azure account
- An Azure app registration with the required permissions
- Python 3.7 or higher
- Pip for installing dependencies

## Setup

### 1. Register Your Application in Azure AD

Before running the script, you'll need to register your app in Azure Active Directory (Azure AD) and set the necessary API permissions:

- Go to [Azure Portal](https://portal.azure.com).
- Navigate to **Azure Active Directory** > **App registrations** > **New registration**.
- Note down the **Application (client) ID**, **Directory (tenant) ID**, and **Client Secret**. These will be required for configuring the script.

### 2. Install Dependencies

Make sure you have `pip` installed, then install the required dependencies:

```bash
pip install -r requirements.txt
```

This will install:

- `Flask` for the web framework.
- `requests` for making HTTP requests.
- `msal` for handling Microsoft OAuth authentication.
- `python-dotenv` for loading environment variables.
- `pytz` for handling timezones.

### 3. Create a `.env` File

Create a `.env` file in the root directory of the project and add the following details:

```bash
CLIENT_ID=<your-client-id>
CLIENT_SECRET=<your-client-secret>
TENANT_ID=<your-tenant-id>
MAILBOX=<your-mailbox-email>
SENDERS=<comma-separated-email-addresses>
```

- `CLIENT_ID`: Your Azure application's client ID.
- `CLIENT_SECRET`: The client secret you generated for your Azure app.
- `TENANT_ID`: The tenant ID for your Azure AD.
- `MAILBOX`: The email address of the mailbox you want to retrieve emails from.
- `SENDERS`: A comma-separated list of senders whose attachments you want to download.

### 4. Run the Application

Start the Flask web server by running:

```bash
python app.py
```

By default, the application will be accessible at `http://localhost:3000`.

### 5. OAuth2 Authentication

When you visit `http://localhost:3000`, you'll be redirected to Microsoft's OAuth2 login page. After logging in and consenting to the requested permissions, you'll be redirected back to the app. The app will then retrieve the emails from the specified mailbox and download the attachments from the specified senders.

### 6. Download Attachments

After successful authentication, the app will process emails received on the current day and download attachments from the specified senders. The attachments will be stored in a folder with the current date and then bundled into a ZIP file, which will be sent as a downloadable file.

## Endpoints

- `/`: Initiates the OAuth2 authentication flow.
- `/token`: Handles the redirect from Microsoft after authentication and retrieves the access token.
- `/download`: Retrieves and downloads email attachments from the specified senders.

## Example Flow

1. Visit `http://localhost:3000/`.
2. Authenticate using Microsoft OAuth2.
3. After authentication, the app fetches emails and attachments from the specified senders.
4. The attachments are packaged into a ZIP file and presented for download.

## Troubleshooting

- **Authorization failed**: Make sure you entered the correct client ID, client secret, tenant ID, and mailbox.
- **No attachments found**: Ensure the specified senders have emails with attachments within the timeframe you are querying.
- **Invalid access token**: Re-authenticate to ensure a valid token is being used.

## License

This project is licensed under the MIT License.
