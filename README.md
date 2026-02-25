# Thrive Admin App (Phase 1)

Standalone admin app with Microsoft Entra sign-in and a secure Clients page.

## Features

- Entra ID sign-in/sign-out (MSAL popup with redirect fallback)
- Authenticated Clients page (`clients.html`)
- Secure backend APIs:
  - `GET /api/clients`
  - `GET /api/clients/:id`
- SharePoint clients source via Graph app-only token
- Optional local fallback client data (`data/clients.json`)

## Run locally

1. Copy `.env.example` to `.env` and set real values.
2. Start server:

```bash
npm start
```

3. Open:

- `http://127.0.0.1:8081/index.html`
- `http://127.0.0.1:8081/clients.html`

## Frontend config

Set values in `frontend-config.js`:

- `tenantId`
- `spaClientId`
- `apiScope` (example: `api://<api-app-id>/client.read`)
- `apiBaseUrl` (empty for same-origin)

## Required backend env vars

- `AZURE_TENANT_ID`
- `AZURE_API_AUDIENCE` or `AZURE_API_CLIENT_ID`
- `AZURE_REQUIRED_SCOPE` (default `client.read`)
- `AZURE_API_CLIENT_SECRET`
- `SHAREPOINT_SITE_URL`
- `SHAREPOINT_CLIENTS_LIST_NAME`

Optional fallback toggles:

- `USE_LOCAL_CLIENTS_FALLBACK=1`
- `ALLOW_LOCAL_CLIENTS_ON_GRAPH_ERROR=1`
- `CLIENTS_DATA_FILE=./data/clients.json`

## Redirect URI notes

For local sign-in, add your SPA redirect in Entra app registration, for example:

- `http://127.0.0.1:8081`

If you use a different host/port, the redirect URI must match exactly.
