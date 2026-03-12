# Thrive Admin App (Phase 1)

Standalone admin app with Microsoft Entra sign-in and a secure Clients page.

## Features

- Entra ID sign-in/sign-out (MSAL popup with redirect fallback)
- Authenticated Clients page (`clients.html`)
- Authenticated Carers page (`carers.html`)
- Authenticated Recruitment page (`recruitment.html`) from SharePoint list (active candidates only)
- Authenticated Tasks page (`task-whiteboard.html`) with draggable pinned cards and category boxes (default)
- Authenticated Tasks (Simple) page (`simple-tasks.html`) with pill view and pinning
- Authenticated Tasks (Advanced) page (`tasks.html`) for full unified To Do + Planner overlay editing
- Authenticated Time Mapping page (`mapping.html`) for run planning
- Authenticated Drive-Time Map page (`drive-time-map.html`) for 20-minute drive-time polygons
- Authenticated Consultant page (`consultant.html`) for anonymisation and DOCX report export
- Secure backend APIs:
  - `GET /api/clients` (original SharePoint/local clients list)
  - `GET /api/clients/:id`
  - `GET /api/onetouch/clients` (OneTouch clients list enriched with SharePoint Xero/consent fields; no carers/visits relationship lookup)
  - `GET /api/clients/reconcile/preview` (admin + care_manager; OneTouch-to-SharePoint reconciliation preview)
  - `POST /api/clients/reconcile/apply` (admin + care_manager; per-record copy/add/update reconcile action)
  - `GET /api/carers`
  - `GET /api/recruitment` (delegated Microsoft Graph token; returns active candidates only)
  - `POST /api/recruitment` (delegated Microsoft Graph token; creates OneTouch carer and writes SharePoint `OnetouchLink`)
  - `POST /api/routes/run`
  - `POST /api/maps/drive-time`
  - `POST /api/maps/geocode-batch`
  - `GET /api/tasks/unified` (delegated Microsoft Graph token)
  - `POST /api/tasks/overlay` (delegated Microsoft Graph token)
  - `POST /api/consultant/report-docx` (admin + consultant; DOCX export from template)
- OneTouch source (`carers/all`, `clients/all`, `visits`) with relationships joined in-app
- Optional local fallback client data (`data/clients.json`)
- Clients reconciliation workflow on `clients.html` treats OneTouch as source of truth and writes changes into SharePoint
  - Reconciliation combines OneTouch multi-contact fields into SharePoint single fields (`email`, `phone`) with dedupe (`; ` for emails, ` / ` for phones)
  - Reconciliation never clears SharePoint fields with blank/null OneTouch values

## Run locally

1. Copy `.env.example` to `.env` and set real values.
2. Start server:

```bash
npm start
```

3. Open:

- `http://127.0.0.1:8081/index.html`
- `http://127.0.0.1:8081/clients.html`
- `http://127.0.0.1:8081/carers.html`
- `http://127.0.0.1:8081/recruitment.html`
- `http://127.0.0.1:8081/task-whiteboard.html`
- `http://127.0.0.1:8081/mapping.html`
- `http://127.0.0.1:8081/drive-time-map.html`

## Frontend config

Set values in `frontend-config.js`:

- `tenantId`
- `spaClientId`
- `apiScope` (example: `api://<api-app-id>/client.read`)
- `graphTaskScopes` (default `User.Read`, `Tasks.ReadWrite`, `Group.Read.All`, `Sites.ReadWrite.All`)
- `apiBaseUrl` (empty for same-origin)

## Required backend env vars

- `AZURE_TENANT_ID`
- `AZURE_API_AUDIENCE` or `AZURE_API_CLIENT_ID`
- `AZURE_REQUIRED_SCOPE` (default `client.read`)
- `SHAREPOINT_SITE_URL`
- `SHAREPOINT_TASK_OVERLAY_LIST_NAME` (optional; default `TaskOverlay`)
- `GRAPH_TOKEN_AUDIENCE` (optional override; defaults include Graph audiences)
- `ONETOUCH_USERNAME`
- `ONETOUCH_PASSWORD`
- `ONETOUCH_CLIENTS_TIMEOUT_MS` (optional; default `12000`)
- `ONETOUCH_CARERS_TIMEOUT_MS` (optional; default `12000`)
- `ONETOUCH_VISITS_TIMEOUT_MS` (optional; default `6000`)
- `ONETOUCH_CARER_DETAIL_CONCURRENCY` (optional; default `4`)

Optional fallback toggles:

- `USE_LOCAL_CLIENTS_FALLBACK=1`
- `ALLOW_LOCAL_CLIENTS_ON_GRAPH_ERROR=1`
- `CLIENTS_DATA_FILE=./data/clients.json`

Google Maps Platform vars (for Time Mapping):

- `GOOGLE_MAPS_API_KEY`
- `GOOGLE_MAPS_REGION` (default `gb`)

Run costing vars (for Time Mapping):

- `MAX_DISTANCE` (miles; applies to home legs only)
- `MAX_TIME` (minutes; used only if `MAX_DISTANCE` is empty)
- `TRAVEL_PAY` (hourly rate for paid travel time)
- `PER_MILE` (rate per paid mile)

## Redirect URI notes

For local sign-in, add your SPA redirect in Entra app registration, for example:

- `http://127.0.0.1:8081`

If you use a different host/port, the redirect URI must match exactly.
