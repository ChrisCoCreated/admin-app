import { FRONTEND_CONFIG } from "./frontend-config.js";

const API_BASE_URL = (FRONTEND_CONFIG.apiBaseUrl || "").replace(/\/+$/, "");

function endpoint(pathname) {
  return API_BASE_URL ? `${API_BASE_URL}${pathname}` : pathname;
}

function buildUrl(pathname, query = {}) {
  const url = new URL(endpoint(pathname), window.location.origin);
  for (const [key, value] of Object.entries(query)) {
    if (value === undefined || value === null || value === "") {
      continue;
    }
    url.searchParams.set(key, String(value));
  }
  return url;
}

async function parseError(response, fallbackLabel) {
  const text = await response.text();
  let payload = null;

  try {
    payload = text ? JSON.parse(text) : null;
  } catch {
    payload = null;
  }

  const structuredError = payload?.error;
  const detail =
    payload?.detail ||
    (typeof structuredError === "string" ? structuredError : structuredError?.message) ||
    text ||
    "Unknown error";
  const error = new Error(`${fallbackLabel} (${response.status}): ${detail}`);
  error.status = response.status;
  error.detail = detail;
  if (structuredError && typeof structuredError === "object") {
    error.code = structuredError.code;
    error.retryable = structuredError.retryable;
    error.correlationId = structuredError.correlationId;
  }
  throw error;
}

export function createDirectoryApi(authController) {
  function resolveScopes(scopeSource) {
    if (Array.isArray(scopeSource) && scopeSource.length > 0) {
      return scopeSource;
    }
    if (typeof scopeSource === "string" && scopeSource.trim()) {
      return [scopeSource.trim()];
    }
    return [FRONTEND_CONFIG.apiScope];
  }

  async function authFetch(pathname, options = {}) {
    const scopes = resolveScopes(options.scopes);
    const token = await authController.acquireToken(scopes);
    const headers = { ...(options.headers || {}) };
    delete options.scopes;
    const response = await fetch(pathname, {
      ...options,
      headers: {
        Accept: "application/json",
        Authorization: `Bearer ${token}`,
        ...headers,
      },
    });

    return response;
  }

  return {
    async getCurrentUser() {
      const response = await authFetch(endpoint("/api/auth/me"));
      if (!response.ok) {
        await parseError(response, "Profile request failed");
      }
      return response.json();
    },

    async listClients(query = {}) {
      const response = await authFetch(buildUrl("/api/clients", query));
      if (!response.ok) {
        await parseError(response, "Clients request failed");
      }
      return response.json();
    },

    async listOneTouchClients(query = {}) {
      const response = await authFetch(buildUrl("/api/onetouch/clients", query));
      if (!response.ok) {
        await parseError(response, "OneTouch clients request failed");
      }
      return response.json();
    },

    async getClientsReconcilePreview() {
      const response = await authFetch(endpoint("/api/clients/reconcile/preview"));
      if (!response.ok) {
        await parseError(response, "Clients reconciliation preview request failed");
      }
      return response.json();
    },

    async applyClientsReconcileAction(payload = {}) {
      const response = await authFetch(endpoint("/api/clients/reconcile/apply"), {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload && typeof payload === "object" ? payload : {}),
      });
      if (!response.ok) {
        await parseError(response, "Clients reconciliation apply request failed");
      }
      return response.json();
    },

    async listCarers(query = {}) {
      const response = await authFetch(buildUrl("/api/carers", query));
      if (!response.ok) {
        await parseError(response, "Carers request failed");
      }
      return response.json();
    },

    async listMarketingPhotos(query = {}) {
      const response = await authFetch(buildUrl("/api/marketing/photos", query));
      if (!response.ok) {
        await parseError(response, "Marketing photos request failed");
      }
      return response.json();
    },

    async getMarketingMedia(query = {}) {
      const response = await authFetch(buildUrl("/api/marketing/media", query));
      if (!response.ok) {
        await parseError(response, "Marketing media request failed");
      }
      return response.json();
    },

    async getUnifiedTasks(query = {}) {
      const response = await authFetch(buildUrl("/api/tasks/unified", query), {
        scopes: FRONTEND_CONFIG.graphTaskScopes,
      });
      if (!response.ok) {
        await parseError(response, "Unified tasks request failed");
      }
      return response.json();
    },

    async getWhiteboardTasks(query = {}) {
      const response = await authFetch(buildUrl("/api/tasks/whiteboard", query), {
        scopes: FRONTEND_CONFIG.graphTaskScopes,
      });
      if (!response.ok) {
        await parseError(response, "Whiteboard tasks request failed");
      }
      return response.json();
    },

    async upsertTaskOverlay({ provider, externalTaskId, patch }) {
      const response = await authFetch(endpoint("/api/tasks/overlay"), {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        scopes: FRONTEND_CONFIG.graphTaskScopes,
        body: JSON.stringify({
          provider,
          externalTaskId,
          patch: patch && typeof patch === "object" ? patch : {},
        }),
      });
      if (!response.ok) {
        await parseError(response, "Task overlay upsert failed");
      }
      return response.json();
    },
  };
}
