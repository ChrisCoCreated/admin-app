# Admin Pseudonymiser Bundle

This bundle contains the deterministic care-note pseudonymising functionality only. It deliberately excludes the desktop shell, Tauri sidecar code, and every Ollama route or UI element.

## Contents

```text
admin-pseudonymiser-bundle/
├── backend/
│   ├── pseudonymiser_core/
│   │   ├── anonymiser.py       # stable placeholder generation
│   │   ├── recognizers.py      # UK social care PII and residual-risk recognisers
│   │   ├── schemas.py          # request/response models, no Ollama fields
│   │   ├── privacy.py          # no raw-note logging guard
│   │   └── router.py           # FastAPI router
│   ├── tests/
│   ├── example_fastapi_app.py
│   └── pyproject.toml
└── frontend/
    ├── PseudonymiserAdminPage.tsx
    ├── pseudonymiserApi.ts
    ├── pseudonymiserTypes.ts
    ├── deanonymise.ts
    ├── pseudonymiser-admin.css
    └── index.ts
```

## Backend API

Mount the bundled router at the same path used by the React page:

```python
from fastapi import FastAPI
from pseudonymiser_core.router import router as pseudonymiser_router

app = FastAPI()
app.include_router(pseudonymiser_router, prefix="/api/pseudonymiser")
```

Endpoints:

- `GET /api/pseudonymiser/health`
- `POST /api/pseudonymiser/pseudonymise` with `{ "text": "..." }`
- `POST /api/pseudonymiser/safety-check` with `{ "text": "..." }`

The pseudonymise response includes `pseudonymised_text`, a local restoration `mapping`, replacement `entities`, aggregate `counts`, and residual-risk `findings`.

## Frontend Usage

Copy the `frontend` files into the target web app and render the page:

```tsx
import { PseudonymiserAdminPage } from "./pseudonymiser";

export default function AdminPseudonymiserRoute() {
  return <PseudonymiserAdminPage apiBaseUrl="/api/pseudonymiser" />;
}
```

The page expects React and `lucide-react`. If the admin app already has its own icon system, replace the icon imports in `PseudonymiserAdminPage.tsx` with local equivalents.

## Privacy Rules

- Keep the `mapping` in browser/session state only unless the product owner explicitly approves storage.
- Never send `mapping` to an external LLM.
- Send only `pseudonymised_text` plus the copied placeholder-preservation instructions to downstream LLM workflows.
- Keep backend access logs disabled or redacted for these routes.
- Use synthetic notes for tests and demos.

## Verification

Backend tests pass from this bundle with:

```bash
cd admin-pseudonymiser-bundle/backend
python -m pytest
```

The frontend files type-check with React 18 and TypeScript using `jsx: react-jsx`.

## Implementation Instructions For The LLM In The New Workspace

You are implementing the bundled care-note pseudonymiser as a page in the admin web app.

1. Inspect the target app first. Identify the frontend framework, route conventions, API/backend stack, styling system, auth/permission patterns, and whether Python services already exist.
2. Import the backend core without adding Ollama, local model checks, Tauri, desktop sidecars, telemetry, or external LLM calls.
3. If the target backend is FastAPI, copy `backend/pseudonymiser_core` into the backend package and include `pseudonymiser_core.router` at `/api/pseudonymiser`.
4. If the target backend is not Python, prefer running this Python core behind an internal service boundary rather than rewriting the recognisers by hand. Preserve the JSON contract exactly if an adapter is needed.
5. Copy the frontend files into the admin app, adapt imports to the local alias/path style, and mount `PseudonymiserAdminPage` as the admin route/page.
6. Keep the UI integrated with the admin shell. Use the host app's nav, auth guard, page title pattern, design tokens, and notification/toast components where obvious.
7. Do not store raw source notes, restored text, or the restoration mapping in application logs, analytics, URLs, or persistent database tables. If product requirements demand persistence, stop and ask for information-governance approval.
8. Keep the copied LLM prompt rules intact: downstream LLMs must preserve bracketed placeholders exactly and must not invent numbering.
9. Add or adapt tests:
   - backend unit tests for pseudonymisation, placeholder stability, residual direct identifiers, and restoration mapping;
   - frontend smoke test for pseudonymise, manual missed-identifier replacement, copy-for-LLM, and restore.
10. Run the target app's formatter, type-check, unit tests, and a browser smoke test of the new admin page before handing back.

Acceptance criteria:

- The admin page can pseudonymise a note using `/api/pseudonymiser/pseudonymise`.
- Placeholder mapping is retained locally for restore, but not sent to external LLMs.
- Manual missed identifiers can be replaced with stable placeholders.
- LLM copy includes the placeholder-preservation instructions and pseudonymised text only.
- Restore mode converts placeholder-bearing LLM output back to the original values using the local mapping.
- There are no Ollama references, routes, buttons, setup instructions, dependencies, or model fields in the imported implementation.
