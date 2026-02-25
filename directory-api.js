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

  const detail = payload?.detail || payload?.error || text || "Unknown error";
  throw new Error(`${fallbackLabel} (${response.status}): ${detail}`);
}

export function createDirectoryApi(authController) {
  async function authFetch(pathname, options = {}) {
    const token = await authController.acquireToken([FRONTEND_CONFIG.apiScope]);
    const response = await fetch(pathname, {
      ...options,
      headers: {
        Accept: "application/json",
        Authorization: `Bearer ${token}`,
        ...(options.headers || {}),
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

    async listCarers(query = {}) {
      const response = await authFetch(buildUrl("/api/carers", query));
      if (!response.ok) {
        await parseError(response, "Carers request failed");
      }
      return response.json();
    },
  };
}
