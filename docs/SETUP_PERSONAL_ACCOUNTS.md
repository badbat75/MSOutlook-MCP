# Setup Guide for Personal Microsoft Accounts

This guide is for setting up the Outlook MCP Server with **Personal Microsoft Accounts** (e.g., @outlook.com, @hotmail.com, @live.com) rather than enterprise/organizational accounts.

## ⚠️ CRITICAL: Account Type Configuration

**IMPORTANT:** When using `OUTLOOK_TENANT_ID=common`, you MUST configure your Azure app registration with:

✅ **"Accounts in any organizational directory and personal Microsoft accounts"**

❌ **NOT** "Personal Microsoft accounts only"

### Why?

The `/common/` endpoint requires the app to have **"All" userAudience** configuration. If you select "Personal Microsoft accounts only", you'll get this error:

```
❌ Authorization Failed
Error: invalid_request
The request is not valid for the application's 'userAudience' configuration.
In order to use /common/ endpoint, the application must not be configured
with 'Consumer' as the user audience.
```

## Key Differences for Personal Accounts

- Use `OUTLOOK_TENANT_ID=common` instead of a specific tenant ID
- No admin consent required (users consent individually)
- Some enterprise-only features may not be available
- App must support "All" account types, not just "Consumer"

## Step-by-Step Setup

### 1. Azure App Registration

1. Go to [Microsoft Entra admin center](https://entra.microsoft.com)
2. Navigate to **Identity → Applications → App registrations → New registration**
3. Configure the registration:

   **Name:** `Outlook MCP Server`

   **Supported account types:**
   - ✅ **MUST SELECT:** "Accounts in any organizational directory and personal Microsoft accounts"
   - ❌ **DO NOT SELECT:** "Personal Microsoft accounts only" (this won't work with `/common/` endpoint!)

   **Redirect URI:**
   - Type: **Web** (not Public client/native)
   - URI: `http://localhost:5000/callback`

4. Click **Register**

**Note:** Even though you're using personal accounts, you must select the "All" option (organizational + personal) to use the `/common/` tenant endpoint.

### 2. Copy Application (Client) ID

After registration, you'll see the Overview page:

1. Copy the **Application (client) ID** (a GUID like `12345678-1234-1234-1234-123456789abc`)
2. Save this for your `.env` file as `OUTLOOK_CLIENT_ID`

### 3. Create Client Secret

1. Go to **Certificates & secrets** (left menu)
2. Click **New client secret**
3. Add a description (e.g., "MCP Server Secret")
4. Choose expiration period (recommend 24 months)
5. Click **Add**
6. **IMPORTANT:** Copy the **Value** (not the Secret ID) immediately - you won't be able to see it again
7. Save this for your `.env` file as `OUTLOOK_CLIENT_SECRET`

### 4. Configure API Permissions

1. Go to **API permissions** (left menu)
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Select **Delegated permissions**
5. Add these permissions:
   - ✅ `Mail.Read`
   - ✅ `Mail.ReadWrite`
   - ✅ `Mail.Send`
   - ✅ `Calendars.Read`
   - ✅ `Calendars.ReadWrite`
   - ✅ `User.Read`
6. Click **Add permissions**

**Note:** Admin consent is NOT required for personal accounts. Users will consent when they first log in.

### 5. Configure Environment Variables

Edit your `.env` file:

```ini
# Copy these from Azure Portal
OUTLOOK_CLIENT_ID=your-application-client-id-from-step-2
OUTLOOK_CLIENT_SECRET=your-client-secret-value-from-step-3
OUTLOOK_TENANT_ID=common
```

**Critical:** For personal accounts, always use `OUTLOOK_TENANT_ID=common`

### 6. Load Environment and Authorize

**Windows:**
```powershell
# Load environment variables and activate venv
. .\scripts\setup-env.ps1

# You should see:
# ================================================================
# Outlook MCP - Environment Setup
# ================================================================
#
# Loading configuration from .env...
# Activating virtual environment...
# Virtual environment activated (venv)
#
# Environment configured:
#   OUTLOOK_CLIENT_ID     = 12345678...
#   OUTLOOK_CLIENT_SECRET = *** (hidden)
#   OUTLOOK_TENANT_ID     = common

# Run authorization (choose one mode)

# Normal mode - Opens browser automatically
python outlook_mcp_auth.py

# Headless mode - For remote/SSH systems
python outlook_mcp_auth.py --no-browser

# Direct mode - Provide auth code directly
python outlook_mcp_auth.py --code 'http://localhost:5000/callback?code=...'
```

**macOS/Linux:**
```bash
# Load environment variables and activate venv
source ./scripts/setup-env.sh

# You should see:
# ================================================================
# Outlook MCP - Environment Setup
# ================================================================
#
# Loading configuration from .env...
# Activating virtual environment...
# Virtual environment activated (venv)
#
# Environment configured:
#   OUTLOOK_CLIENT_ID     = 12345678...
#   OUTLOOK_CLIENT_SECRET = *** (hidden)
#   OUTLOOK_TENANT_ID     = common

# Run authorization (choose one mode)

# Normal mode - Opens browser automatically
python outlook_mcp_auth.py

# Headless mode - For remote/SSH systems
python outlook_mcp_auth.py --no-browser

# Direct mode - Provide auth code directly
python outlook_mcp_auth.py --code 'http://localhost:5000/callback?code=...'
```

### 7. Sign In with Your Personal Account

**Normal Mode Workflow:**

The authorization script will:

1. Open your default browser
2. Navigate to Microsoft login page
3. Ask you to sign in with your **personal Microsoft account**
   - @outlook.com
   - @hotmail.com
   - @live.com
   - etc.
4. Show the permissions the app is requesting
5. Ask you to consent to these permissions
6. Redirect to `http://localhost:5000/callback`
7. Save the OAuth tokens to `~/.outlook_mcp_token_cache.json`

**💡 TIP:** If the callback doesn't work (firewall, network issues), press **Ctrl+C** and paste the callback URL manually from your browser's address bar.

**Headless Mode Workflow (for remote/SSH systems):**

The script will:

1. Display an authorization URL
2. Prompt you to copy this URL and open it on ANY device (phone, laptop, etc.)
3. After signing in, the browser redirects to `http://localhost:5000/callback?code=...`
4. Copy this FULL URL from the browser address bar
5. Paste it into the terminal prompt
6. Save the OAuth tokens to `~/.outlook_mcp_token_cache.json`

**Success Message:**

You should see:

```
✅ Authentication successful!
   Token cache saved to: C:\Users\YourUser\.outlook_mcp_token_cache.json
   Scopes granted: Mail.Read, Mail.ReadWrite, ...

You can now start the MCP server:
   python outlook_mcp_server.py
```

### 8. Test the Server

```powershell
python outlook_mcp_server.py
```

The server should start without errors. You'll see output like:

```
INFO:outlook_mcp:Outlook MCP Server starting...
INFO:outlook_mcp:Loaded cached tokens
```

Press `Ctrl+C` to stop the server when testing is complete.

## Troubleshooting

### Error: "Cannot open browser on remote/SSH system"

**Cause:** You're on a headless server or remote SSH session without GUI.

**Solution:**

Use headless mode:

**Windows:**
```powershell
. .\scripts\setup-env.ps1
python outlook_mcp_auth.py --no-browser
```

**macOS/Linux:**
```bash
source ./scripts/setup-env.sh
python outlook_mcp_auth.py --no-browser
```

**Steps:**
1. Script displays authorization URL
2. Copy URL and open on ANY device with a browser (phone, laptop, etc.)
3. Sign in with your Microsoft account
4. After sign-in, browser redirects to `http://localhost:5000/callback?code=...`
5. Copy the FULL URL from browser address bar
6. Paste it into the terminal prompt
7. Script completes authorization and saves tokens

**Alternative:** If callback doesn't work in normal mode, press Ctrl+C when prompted and paste the URL manually.

---

### Error: "Browser opens but callback never arrives"

**Cause:** Firewall, network issues, or browser security settings blocking `localhost:5000` callback.

**Solution:**

Don't restart! Just press **Ctrl+C** in the running auth script and paste the URL manually:

1. Script is waiting: `Waiting for authorization callback on http://localhost:5000 ...`
2. Press **Ctrl+C**
3. Script prompts: `Paste Callback URL:`
4. Look at your browser address bar after signing in - it should show `http://localhost:5000/callback?code=...`
5. Copy the FULL URL from the address bar
6. Paste it into the terminal
7. Script completes authorization

**💡 TIP:** The script shows this hint while waiting: "If the callback doesn't work, press Ctrl+C to paste manually"

---

### Error: "invalid_request - not valid for the application's 'userAudience' configuration" ⚠️ MOST COMMON

**Full Error Message:**
```
❌ Authorization Failed
Error: invalid_request

The request is not valid for the application's 'userAudience' configuration.
In order to use /common/ endpoint, the application must not be configured
with 'Consumer' as the user audience. The userAudience should be configured
with 'All' to use /common/ endpoint.
```

**Cause:** Your app registration is set to "Personal Microsoft accounts only" (Consumer), but the `/common/` endpoint requires "All" (organizational + personal).

**Solution - Option 1: Update Existing App (Easier)**

1. Go to [Azure Portal](https://portal.azure.com) → **App registrations** → Your app
2. Click **Authentication** in the left menu
3. Under **Supported account types**, change to:
   - ✅ **"Accounts in any organizational directory and personal Microsoft accounts"**
4. Click **Save**
5. Run the setup script then re-authorize:
   - **Windows:** `. .\scripts\setup-env.ps1` then `python outlook_mcp_auth.py`
   - **macOS/Linux:** `source ./scripts/setup-env.sh` then `python outlook_mcp_auth.py`

**Solution - Option 2: Create New App**

1. Delete the current app registration (or keep it for reference)
2. Follow the setup guide from Step 1, making sure to select:
   - ✅ "Accounts in any organizational directory and personal Microsoft accounts"
3. Update your `.env` file with the new Client ID and Secret
4. Run the setup script then re-authorize:
   - **Windows:** `. .\scripts\setup-env.ps1` then `python outlook_mcp_auth.py`
   - **macOS/Linux:** `source ./scripts/setup-env.sh` then `python outlook_mcp_auth.py`

---

### Error: "AADSTS50020: User account from identity provider does not exist in tenant"

**Cause:** You're using a personal account but `OUTLOOK_TENANT_ID` is set to a specific tenant ID instead of "common".

**Solution:**
```ini
# In your .env file, change to:
OUTLOOK_TENANT_ID=common
```

Then reload and re-run:
- **Windows:** `. .\scripts\setup-env.ps1` and `python outlook_mcp_auth.py`
- **macOS/Linux:** `source ./scripts/setup-env.sh` and `python outlook_mcp_auth.py`

### Error: "AADSTS7000218: The request body must contain the following parameter: 'client_assertion' or 'client_secret'"

**Cause:** The client secret is missing or incorrect in your `.env` file.

**Solution:**
1. Go to Azure Portal → Your app → Certificates & secrets
2. Create a new client secret (the old one may have expired)
3. Copy the **Value** (not Secret ID)
4. Update `OUTLOOK_CLIENT_SECRET` in `.env`
5. Reload:
   - **Windows:** `. .\scripts\setup-env.ps1`
   - **macOS/Linux:** `source ./scripts/setup-env.sh`

### Error: "AADSTS700016: Application with identifier '...' was not found in the directory"

**Cause:** The client ID is incorrect.

**Solution:**
1. Go to Azure Portal → App registrations → Your app → Overview
2. Copy the correct **Application (client) ID**
3. Update `OUTLOOK_CLIENT_ID` in `.env`
4. Reload:
   - **Windows:** `. .\scripts\setup-env.ps1`
   - **macOS/Linux:** `source ./scripts/setup-env.sh`

### Error: "AADSTS50011: The redirect URI '...' specified in the request does not match"

**Cause:** The redirect URI in the app registration doesn't match what the code is using.

**Solution:**
1. Go to Azure Portal → Your app → Authentication
2. Ensure there's a **Web** redirect URI (not Public client)
3. Ensure it's exactly: `http://localhost:5000/callback`
4. Save changes and try again: `python outlook_mcp_auth.py`

### Error: "401 Unauthorized" when running the server

**Cause:** Token has expired or was revoked.

**Solution:**
```powershell
# Re-authorize
python outlook_mcp_auth.py
```

### Error: "403 Forbidden" when trying to access mail/calendar

**Cause:** Missing API permissions.

**Solution:**
1. Go to Azure Portal → Your app → API permissions
2. Verify all required permissions are listed
3. For personal accounts, no admin consent is needed
4. Re-authorize: `python outlook_mcp_auth.py`

## Security Notes

- **Token Storage:** OAuth tokens are stored locally in `~/.outlook_mcp_token_cache.json`
- **Token Expiration:** Tokens are automatically refreshed by MSAL
- **Secret Expiration:** Client secrets expire (check Azure Portal for expiration date)
- **Revoking Access:** Go to [Microsoft Account Privacy Settings](https://account.microsoft.com/privacy) → App permissions → Remove the app

## Next Steps

Once setup is complete:

1. **Configure Claude Desktop** - See [QUICKSTART.md](QUICKSTART.md#5-configure-claude-desktop)
2. **Test with Claude** - Try commands like "Show me my unread emails"
3. **Daily Usage** - Just run the setup script to load environment:
   - **Windows:** `. .\scripts\setup-env.ps1`
   - **macOS/Linux:** `source ./scripts/setup-env.sh`

## Reference

### Correct Configuration Summary

| Setting | Required Value | ❌ Wrong Value |
|---------|---------------|---------------|
| **Azure: Supported account types** | ✅ "Accounts in any organizational directory and personal Microsoft accounts" | ❌ "Personal Microsoft accounts only" |
| **OUTLOOK_TENANT_ID** | `common` | Specific tenant GUID |
| **Redirect URI Type** | Web | Public client/native |
| **Redirect URI** | `http://localhost:5000/callback` | Any other URL |
| **Account Types Supported** | @outlook.com, @hotmail.com, @live.com, etc. | N/A |
| **Admin Consent** | Not required | N/A |

### Why "All" Account Types?

Even though you're only using personal accounts, you must select "Accounts in any organizational directory and personal Microsoft accounts" because:

- The `/common/` endpoint requires `userAudience: "All"`
- "Personal Microsoft accounts only" sets `userAudience: "Consumer"`
- Consumer userAudience is incompatible with `/common/` endpoint
- This is a Microsoft platform limitation, not a bug in this code

## Additional Resources

- [Microsoft Graph API Documentation](https://learn.microsoft.com/en-us/graph/)
- [MSAL Python Documentation](https://msal-python.readthedocs.io/)
- [Azure App Registration Guide](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-register-app)
