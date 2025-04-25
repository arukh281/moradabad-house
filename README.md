# Moradabad House Management System

This repository contains various tools for managing Moradabad House operations, including WhatsApp notifications, statement uploads, and invoice management.

## Setup Instructions

1. Clone this repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Create a `.env` file in the root directory with the following variables:
   ```
   # Flask
   FLASK_SECRET_KEY=your_secret_key_here

   # Twilio
   TWILIO_ACCOUNT_SID=your_account_sid_here
   TWILIO_AUTH_TOKEN=your_auth_token_here

   # Google Sheets
   GOOGLE_SHEETS_CREDENTIALS_FILE=path_to_your_credentials.json
   ```

4. For Google Sheets integration:
   - Create a service account in Google Cloud Console
   - Download the credentials JSON file
   - Place it in a secure location outside the repository
   - Update the `GOOGLE_SHEETS_CREDENTIALS_FILE` path in your `.env` file

## Security Note

Never commit the following files to version control:
- `.env` file
- Any JSON credential files
- Any files containing API keys or secrets

These files are already included in `.gitignore` to prevent accidental commits.

## Project Structure

- `whatsapp bot/`: WhatsApp notification system
- `statement upload/`: Tools for uploading and managing statements
- `invoice upload/`: Invoice management system

## License

[Add your license information here] 