# Clippy

CLI for Microsoft 365/OWA using Bearer token authentication.

## Installation

```bash
bun install
```

## Authentication

Clippy uses Bearer JWT tokens to authenticate with OWA.

### Option 1: Interactive login (recommended)

Run the login command with `--interactive` to open a browser and automatically extract the token:

```bash
bun run src/cli.ts login --interactive
```

This will:
1. Open a Chrome browser window
2. Navigate to outlook.office.com
3. Capture the Bearer token from network requests
4. Validate and use the token

If you're not already logged in, complete the login in the browser window.

### Option 2: Manual token extraction

1. Open https://outlook.office.com/mail in Chrome/Edge
2. Open DevTools (F12) → Network tab
3. Filter by `service.svc`
4. Click any request and look at Request Headers
5. Copy the `Authorization: Bearer eyJ...` value (without "Bearer ")
6. Set as environment variable:
   ```bash
   export CLIPPY_TOKEN="eyJ0eXAiOiJKV1Q..."
   ```

### Token Details

- **Audience:** `https://outlook.office.com`
- **Lifetime:** ~73 minutes (tokens expire and need re-extraction)
- **Scopes:** Mail.ReadWrite, Calendars.ReadWrite, Contacts.ReadWrite, etc.

## Usage

```bash
# Interactive login (opens browser)
bun run src/cli.ts login --interactive

# Check login status (requires CLIPPY_TOKEN or prior interactive login)
bun run src/cli.ts login

# View account info
bun run src/cli.ts whoami
```

## API

The CLI uses the Outlook REST API v2.0:
- Base URL: `https://outlook.office.com/api/v2.0/me/`
- Auth: `Authorization: Bearer <token>` header
