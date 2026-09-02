# Thrive Admin App (Phase 1)

Standalone admin app with Microsoft Entra sign-in and a secure Clients page.

## Features

- Entra ID sign-in/sign-out (MSAL popup with redirect fallback)
- Authenticated Clients page (`clients.html`)
- Permission-gated Client Contacts page (`client-contacts.html`) with names and emails for active clients
- Authenticated Carers page (`carers.html`)
- Authenticated Recruitment page (`recruitment.html`) from SharePoint list (active candidates only)
- Authenticated Tasks page (`task-whiteboard.html`) with draggable pinned cards and category boxes (default)
- Authenticated Tasks (Simple) page (`simple-tasks.html`) with pill view and pinning
- Authenticated Tasks (Advanced) page (`tasks.html`) for full unified To Do + Planner overlay editing
- Authenticated Time Mapping page (`mapping.html`) for run planning
- Authenticated Drive-Time Map page (`drive-time-map.html`) for 20-minute drive-time polygons
- Secure backend APIs:
  - `GET /api/sharepoint/clients` (original SharePoint/local clients list)
  - `GET /api/sharepoint/clients/:id`
  - `GET /api/onetouch/clients` (OneTouch clients list enriched with SharePoint Xero/consent fields; no carers/visits relationship lookup)
  - `GET /api/client-contacts` (permission-gated active client names and email addresses)
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
  - `POST /api/tasks/assign` (API auth + Graph app token; create To Do tasks for arbitrary users or from SharePoint Task Sets)
  - `POST /api/tasks/create-batch` (API auth + Graph app token; Planner test batch creation from direct payloads or SharePoint Task Sets)
  - `GET /api/tasks/task-sets` (API auth; live SharePoint Task Set rows for preview/testing)
  - `POST /api/consultant/report-docx` (admin + consultant; DOCX export from template)
  - `POST /api/ai/chat` (API auth; provider-switched chat completions via DeepSeek or Azure OpenAI)
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

Optional checks:

```bash
npm test
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
- `AZURE_API_CLIENT_SECRET` for app-only Graph calls
- `CRON_SECRET` for Vercel Cron authorization
- At least one app access env var such as `ACCESS_FULL_EMAILS` or `ACCESS_DIRECTOR_EMAILS`
- `SHAREPOINT_SITE_URL`
- `SHAREPOINT_ENQUIRIES_SITE_URL` (optional; defaults to ThriveCalls)
- `SHAREPOINT_ENQUIRIES_LIST_NAME` (optional; default `Enquiries Log`)
- `SHAREPOINT_ENQUIRIES_LIST_WEB_URL` (optional; used for reliable list resolution and email links)
- `ENQUIRY_REMINDER_FROM_EMAIL` for the mailbox that sends enquiry reminder emails
- `ENQUIRY_REMINDER_RECIPIENT_OVERRIDE` for trial delivery; set to `chris@planwithcare.co.uk`
- `SHAREPOINT_TASK_OVERLAY_LIST_NAME` (optional; default `TaskOverlay`)
- `SHAREPOINT_TASK_SETS_SITE_URL` (optional; defaults to `SHAREPOINT_SITE_URL`)
- `SHAREPOINT_TASK_SETS_LIST_NAME` (optional; default `Actions for Task Sets`)
- `PLANNER_TEST_PLAN_ID` (used by the Planner test batch creator)
- `PLANNER_TEST_BUCKET_NAME` (used by the Planner test batch creator)
- `GRAPH_TOKEN_AUDIENCE` (optional override; defaults include Graph audiences)
- `ONETOUCH_USERNAME`
- `ONETOUCH_PASSWORD`
- `ONETOUCH_CLIENTS_TIMEOUT_MS` (optional; default `12000`)
- `ONETOUCH_CARERS_TIMEOUT_MS` (optional; default `12000`)
- `ONETOUCH_VISITS_TIMEOUT_MS` (optional; default `6000`)
- `ONETOUCH_CARER_DETAIL_CONCURRENCY` (optional; default `4`)
- `AI_PROVIDER` (optional; default `deepseek`)
- `DEEPSEEK_API_KEY` when `AI_PROVIDER=deepseek`
- `DEEPSEEK_MODEL` (optional; default `deepseek-v4-flash`)
- `AZURE_OPENAI_ENDPOINT` when `AI_PROVIDER=azure_openai`
- `AZURE_OPENAI_API_KEY` when `AI_PROVIDER=azure_openai`
- `AZURE_OPENAI_API_VERSION` when `AI_PROVIDER=azure_openai`
- `AZURE_OPENAI_DEPLOYMENT_NAME` when `AI_PROVIDER=azure_openai` as a single deployment fallback
- `AZURE_OPENAI_DEPLOYMENT_PRIMARY` when `AI_PROVIDER=azure_openai` for primary/pro requests
- `AZURE_OPENAI_DEPLOYMENT_FAST` when `AI_PROVIDER=azure_openai` for fast/flash requests

Optional fallback toggles:

- `USE_LOCAL_CLIENTS_FALLBACK=1`
- `ALLOW_LOCAL_CLIENTS_ON_GRAPH_ERROR=1`
- `CLIENTS_DATA_FILE=./data/clients.json`

For `/api/tasks/create-batch`, the Entra app needs Graph application permissions that can create Planner tasks and read the SharePoint Task Set list. In practice this means `Tasks.ReadWrite.All`, site read access such as `Sites.Read.All`, and enough directory/user read access to resolve assignee identities.

Google Maps Platform vars (for Time Mapping):

- `GOOGLE_MAPS_API_KEY`
- `GOOGLE_MAPS_REGION` (default `gb`)

## Scheduled enquiry reminders

`vercel.json` includes a weekly Vercel Cron entry for `/api/cron/enquiry-reminders` at `0 7 * * 1`, which is Monday 8am UK time during BST. If exact winter 8am UK delivery matters, change this to `0 8 * * 1` when the UK is on GMT.

The route requires `Authorization: Bearer ${CRON_SECRET}`. It emails active enquiries not modified in the last 7 days every week, and includes on-hold enquiries only on the first Monday of each month. For the trial, keep `ENQUIRY_REMINDER_RECIPIENT_OVERRIDE=chris@planwithcare.co.uk` so all reminders go to Chris rather than enquiry owners.

The Entra app used for Graph app-only calls must be able to read the SharePoint enquiries list and send mail from `ENQUIRY_REMINDER_FROM_EMAIL`, typically using `Sites.Read.All` or equivalent site access plus `Mail.Send`.

Run costing vars (for Time Mapping):

- `MAX_DISTANCE` (miles; applies to home legs only)
- `MAX_TIME` (minutes; used only if `MAX_DISTANCE` is empty)
- `TRAVEL_PAY` (hourly rate for paid travel time)
- `PER_MILE` (rate per paid mile)

## AI provider notes

The app defaults to DeepSeek. To switch the existing AI routes and the Wellbeing Intake AI extraction flow to Azure OpenAI, set:

```bash
AI_PROVIDER=azure_openai
AZURE_OPENAI_ENDPOINT=https://your-resource.openai.azure.com
AZURE_OPENAI_API_KEY=your-key
AZURE_OPENAI_API_VERSION=2024-10-21
AZURE_OPENAI_DEPLOYMENT_PRIMARY=ai-primary-prod
AZURE_OPENAI_DEPLOYMENT_FAST=ai-fast-prod
```

When Azure OpenAI is enabled, the server can map the existing app model choices onto Azure deployment names. `deepseek-v4-pro` routes to `AZURE_OPENAI_DEPLOYMENT_PRIMARY` and `deepseek-v4-flash` routes to `AZURE_OPENAI_DEPLOYMENT_FAST`. `AZURE_OPENAI_DEPLOYMENT_NAME` remains available as a single-deployment fallback.

## App access roles

Access is environment-driven only. `data/authorized-users.json` is not used for runtime authorization.

Set comma-separated email lists with these env vars:

- `ACCESS_FULL_EMAILS` for `admin`
- `ACCESS_DIRECTOR_EMAILS` for `director`
- `ACCESS_MARKETING_EMAILS` for `marketing`
- `ACCESS_PHOTO_LAYOUT_EMAILS` for `photo_layout`
- `ACCESS_FINANCE_EMAILS` for `finance`
- `ACCESS_TIME_EMAILS` for time-mapping roles
- `ACCESS_HR_EMAILS` for HR roles
- `ACCESS_CLIENTS_EMAILS` for client roles
- `ACCESS_ENQUIRIES_EMAILS` for enquiries-only access
- `ACCESS_CONSULTANT_EMAILS` for `consultant`

Recruitment access is granted by these existing env assignments:

- `ACCESS_FULL_EMAILS`
- `ACCESS_HR_EMAILS`
- `ACCESS_HR_EMAILS` + `ACCESS_CLIENTS_EMAILS`
- `ACCESS_TIME_EMAILS` + `ACCESS_HR_EMAILS`
- `ACCESS_TIME_EMAILS` + `ACCESS_HR_EMAILS` + `ACCESS_CLIENTS_EMAILS`

## Redirect URI notes

For local sign-in, add your SPA redirect in Entra app registration, for example:

- `http://127.0.0.1:8081`

If you use a different host/port, the redirect URI must match exactly.
