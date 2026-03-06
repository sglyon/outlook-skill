# Outlook Manual Setup Guide

Use this guide to register an app in Microsoft Entra ID and authenticate the Python CLI.

## Prerequisites

- Python 3.11+
- [uv](https://docs.astral.sh/uv/) installed
- Microsoft account (Outlook.com, Hotmail, Live, or Microsoft 365)
- Access to [Microsoft Entra ID](https://entra.microsoft.com)

## Step 1: Create App Registration in Entra ID

1. Go to https://entra.microsoft.com
2. Navigate to **Identity** -> **Applications** -> **App registrations**
3. Click **+ New registration**
4. Configure:
   - **Name:** `Outlook-CLI` (or any name)
   - **Supported account types:** "Accounts in any organizational directory and personal Microsoft accounts"
   - **Redirect URI:** Select **Mobile and desktop applications**, URI = `https://login.microsoftonline.com/common/oauth2/nativeclient`
5. Click **Register**

## Step 2: Note the Client ID

After registration:
1. On the app overview page, copy the **Application (client) ID** -- this is your `CLIENT_ID`
2. No client secret is needed. This is a public client application using device code flow.

## Step 3: Configure API Permissions

1. Go to **API permissions** in the left menu
2. Click **+ Add a permission**
3. Select **Microsoft Graph** -> **Delegated permissions**
4. Add these permissions:
   - `Mail.ReadWrite` - Read and write mail
   - `Mail.Send` - Send mail
   - `Calendars.ReadWrite` - Read and write calendar
   - `User.Read` - Read user profile
5. Click **Add permissions**

Note: `offline_access` is requested automatically during authentication.

## Step 4: Run Setup

```bash
uv run outlook.py setup
```

The setup command will:
1. Prompt for your Application (client) ID from Step 2
2. Initiate device code flow -- open the URL shown in your browser and enter the displayed code
3. After authentication, MSAL stores tokens in `~/.outlook-mcp/default/msal_cache.json`

For additional accounts:
```bash
uv run outlook.py --account work setup
```

## Step 5: Verify Setup

```bash
uv run outlook.py token test
```

You should see a success message confirming access to Microsoft Graph.

## Troubleshooting

### "AADSTS700016: Application not found"
- Double-check the client_id is correct
- Ensure you selected "Accounts in any organizational directory and personal Microsoft accounts"

### "AADSTS65001: User hasn't consented"
- Re-run `uv run outlook.py setup` to go through the consent flow again
- Make sure you click "Accept" on the consent screen

### "Token expired"
- MSAL handles refresh automatically in most cases
- Run `uv run outlook.py token refresh` to force a token refresh

### Work/School Account Issues
- Your organization may require admin consent for the app permissions
- Contact your IT admin or use a personal Microsoft account
