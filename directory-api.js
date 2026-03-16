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

    async getPerformanceScorecard(query = {}) {
      const response = await authFetch(buildUrl("/api/scorecard", query));
      if (!response.ok) {
        await parseError(response, "Performance scorecard request failed");
      }
      return response.json();
    },

    async upsertPerformanceScorecard(payload = {}) {
      const response = await authFetch(endpoint("/api/scorecard"), {
        method: "PUT",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload && typeof payload === "object" ? payload : {}),
      });
      if (!response.ok) {
        await parseError(response, "Performance scorecard save failed");
      }
      return response.json();
    },

    async createPerformanceScorecardDefinition(payload = {}) {
      const response = await authFetch(endpoint("/api/scorecard/definitions"), {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload && typeof payload === "object" ? payload : {}),
      });
      if (!response.ok) {
        await parseError(response, "Performance scorecard definition create failed");
      }
      return response.json();
    },

    async listPerformanceScorecardDefinitions(query = {}) {
      const response = await authFetch(buildUrl("/api/scorecard/definitions", query));
      if (!response.ok) {
        await parseError(response, "Performance scorecard definitions request failed");
      }
      return response.json();
    },

    async updatePerformanceScorecardDefinition(payload = {}) {
      const response = await authFetch(endpoint("/api/scorecard/definitions"), {
        method: "PATCH",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload && typeof payload === "object" ? payload : {}),
      });
      if (!response.ok) {
        await parseError(response, "Performance scorecard definition update failed");
      }
      return response.json();
    },

    async listAgendas(query = {}) {
      const response = await authFetch(buildUrl("/api/agendas", query));
      if (!response.ok) {
        await parseError(response, "Agendas request failed");
      }
      return response.json();
    },

    async getAgendaDetail(query = {}) {
      const response = await authFetch(buildUrl("/api/agendas/detail", query));
      if (!response.ok) {
        await parseError(response, "Agenda detail request failed");
      }
      return response.json();
    },

    async createAgenda(payload = {}) {
      const response = await authFetch(endpoint("/api/agendas"), {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload && typeof payload === "object" ? payload : {}),
      });
      if (!response.ok) {
        await parseError(response, "Agenda create failed");
      }
      return response.json();
    },

    async updateAgenda(payload = {}) {
      const response = await authFetch(endpoint("/api/agendas"), {
        method: "PATCH",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload && typeof payload === "object" ? payload : {}),
      });
      if (!response.ok) {
        await parseError(response, "Agenda update failed");
      }
      return response.json();
    },

    async deleteAgenda(payload = {}) {
      const response = await authFetch(buildUrl("/api/agendas", payload), {
        method: "DELETE",
      });
      if (!response.ok) {
        await parseError(response, "Agenda delete failed");
      }
      return response.json();
    },

    async createAgendaItem(payload = {}) {
      const response = await authFetch(endpoint("/api/agendas/items"), {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload && typeof payload === "object" ? payload : {}),
      });
      if (!response.ok) {
        await parseError(response, "Agenda item create failed");
      }
      return response.json();
    },

    async updateAgendaItem(payload = {}) {
      const response = await authFetch(endpoint("/api/agendas/items"), {
        method: "PATCH",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload && typeof payload === "object" ? payload : {}),
      });
      if (!response.ok) {
        await parseError(response, "Agenda item update failed");
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

    async getOneTouchTags() {
      const response = await authFetch(endpoint("/api/onetouch/tags"));
      if (!response.ok) {
        await parseError(response, "OneTouch tags request failed");
      }
      return response.json();
    },

    async exportConsultantReportDocx(payload = {}) {
      const response = await authFetch(endpoint("/api/consultant/report-docx"), {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload && typeof payload === "object" ? payload : {}),
      });
      if (!response.ok) {
        await parseError(response, "Consultant report export failed");
      }
      return response.blob();
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

    async applyBulkClientTag(payload = {}) {
      const response = await authFetch(endpoint("/api/clients/tags/bulk"), {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload && typeof payload === "object" ? payload : {}),
      });
      if (!response.ok) {
        await parseError(response, "Bulk client tag request failed");
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

    async applyBulkCarerTag(payload = {}) {
      const response = await authFetch(endpoint("/api/carers/tags/bulk"), {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload && typeof payload === "object" ? payload : {}),
      });
      if (!response.ok) {
        await parseError(response, "Bulk carer tag request failed");
      }
      return response.json();
    },

    async listTimesheets(query = {}) {
      const response = await authFetch(buildUrl("/api/timesheets", query));
      if (!response.ok) {
        await parseError(response, "Timesheets request failed");
      }
      return response.json();
    },

    async listRecruitment(query = {}) {
      const response = await authFetch(buildUrl("/api/recruitment", query), {
        scopes: FRONTEND_CONFIG.graphTaskScopes,
      });
      if (!response.ok) {
        await parseError(response, "Recruitment request failed");
      }
      return response.json();
    },

    async addRecruitmentCandidateToOneTouch(payload = {}) {
      const response = await authFetch(endpoint("/api/recruitment"), {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        scopes: FRONTEND_CONFIG.graphTaskScopes,
        body: JSON.stringify(payload && typeof payload === "object" ? payload : {}),
      });
      if (!response.ok) {
        await parseError(response, "Recruitment create request failed");
      }
      return response.json();
    },

    async updateRecruitmentStatus(payload = {}) {
      const response = await authFetch(endpoint("/api/recruitment/status"), {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        scopes: FRONTEND_CONFIG.graphTaskScopes,
        body: JSON.stringify(payload && typeof payload === "object" ? payload : {}),
      });
      if (!response.ok) {
        await parseError(response, "Recruitment status update request failed");
      }
      return response.json();
    },

    async getRecruitmentOneTouchOptions() {
      const response = await authFetch(endpoint("/api/recruitment/onetouch-options"), {
        scopes: FRONTEND_CONFIG.graphTaskScopes,
      });
      if (!response.ok) {
        await parseError(response, "Recruitment OneTouch options request failed");
      }
      return response.json();
    },

    async previewRecruitmentImport(payload = {}) {
      const response = await authFetch(endpoint("/api/recruitment/import"), {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        scopes: FRONTEND_CONFIG.graphTaskScopes,
        body: JSON.stringify({
          ...(payload && typeof payload === "object" ? payload : {}),
          dryRun: true,
        }),
      });
      if (!response.ok) {
        await parseError(response, "Recruitment import preview request failed");
      }
      return response.json();
    },

    async runRecruitmentImport(payload = {}) {
      const response = await authFetch(endpoint("/api/recruitment/import"), {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        scopes: FRONTEND_CONFIG.graphTaskScopes,
        body: JSON.stringify({
          ...(payload && typeof payload === "object" ? payload : {}),
          dryRun: false,
        }),
      });
      if (!response.ok) {
        await parseError(response, "Recruitment import request failed");
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

    async syncWhiteboardTasks(payload = {}) {
      const response = await authFetch(endpoint("/api/tasks/whiteboard-sync"), {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        scopes: FRONTEND_CONFIG.graphTaskScopes,
        body: JSON.stringify(payload && typeof payload === "object" ? payload : {}),
      });
      if (!response.ok) {
        await parseError(response, "Whiteboard sync request failed");
      }
      return response.json();
    },

    async createTask(payload = {}) {
      const response = await authFetch(endpoint("/api/tasks/create"), {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        scopes: FRONTEND_CONFIG.graphTaskScopes,
        body: JSON.stringify(payload && typeof payload === "object" ? payload : {}),
      });
      if (!response.ok) {
        await parseError(response, "Task create failed");
      }
      return response.json();
    },

    async updateTask(payload = {}) {
      const response = await authFetch(endpoint("/api/tasks/update"), {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        scopes: FRONTEND_CONFIG.graphTaskScopes,
        body: JSON.stringify(payload && typeof payload === "object" ? payload : {}),
      });
      if (!response.ok) {
        await parseError(response, "Task update failed");
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
