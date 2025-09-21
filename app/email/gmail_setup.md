# Gmail Setup for PPT Email Sending

To enable email sending functionality, you need to set up Gmail credentials.

## Required Environment Variables

Add these variables to your `.env` file:

```
GMAIL_EMAIL=your-email@gmail.com
GMAIL_APP_PASSWORD=your-app-specific-password
```

## Setting up Gmail App Password

1. **Enable 2-Factor Authentication** on your Gmail account (required for app passwords)

2. **Generate App Password:**
   - Go to [Google Account Settings](https://myaccount.google.com/)
   - Navigate to Security → 2-Step Verification
   - Scroll down to "App passwords"
   - Select "Mail" as the app and choose your device
   - Copy the generated 16-character password

3. **Update Environment Variables:**
   - Add your Gmail address to `GMAIL_EMAIL`
   - Add the app password to `GMAIL_APP_PASSWORD`

## Example .env file:
```
# OpenAI/Anthropic API Keys
OPENAI_API_KEY=your-openai-key
ANTHROPIC_API_KEY=your-anthropic-key

# Gmail Configuration
GMAIL_EMAIL=your-email@gmail.com
GMAIL_APP_PASSWORD=abcd efgh ijkl mnop
```

## Security Notes

- **Never share your app password** - treat it like a regular password
- The app password is specific to this application
- You can revoke it anytime from your Google Account settings
- Use environment variables to keep credentials secure

## Testing Email Functionality

Once configured, you can test email sending by:

1. Running the PPT generator with an email recipient
2. The system will automatically send the generated presentation
3. Check the console output for email sending status

## Troubleshooting

- **Authentication Error:** Verify app password is correct
- **SMTP Error:** Check internet connection and Gmail settings
- **File Not Found:** Ensure presentation was generated successfully before email sending
