import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js";

const searchInput = document.getElementById("searchInput");
const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const clientStatusFilters = document.getElementById("clientStatusFilters");
const clientsTableBody = document.getElementById("clientsTableBody");
const emptyState = document.getElementById("emptyState");
const warningState = document.getElementById("warningState");
const detailRoot = document.getElementById("clientDetail");
const linkedCarersList = document.getElementById("linkedCarersList");

const detailFields = {
  id: detailRoot?.querySelector('[data-field="id"]'),
  name: detailRoot?.querySelector('[data-field="name"]'),
  postcode: detailRoot?.querySelector('[data-field="postcode"]'),
  email: detailRoot?.querySelector('[data-field="email"]'),
  phone: detailRoot?.querySelector('[data-field="phone"]'),
  status: detailRoot?.querySelector('[data-field="status"]'),
  tags: detailRoot?.querySelector('[data-field="tags"]'),
  visitCount: detailRoot?.querySelector('[data-field="visitCount"]'),
  carerCount: detailRoot?.querySelector('[data-field="carerCount"]'),
  lastVisitAt: detailRoot?.querySelector('[data-field="lastVisitAt"]'),
};

const DEFAULT_STATUS_FILTERS = new Set(["active", "pending"]);

let allClients = [];
let selectedClientId = "";
let selectedClientStatuses = new Set(DEFAULT_STATUS_FILTERS);
let account = null;

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
  onSignedIn: (signedInAccount) => {
    account = signedInAccount;
  },
  onSignedOut: () => {
    account = null;
  },
});

const directoryApi = createDirectoryApi(authController);

function redirectToUnauthorized(pageKey) {
  const page = encodeURIComponent(String(pageKey || "clients").trim().toLowerCase());
  window.location.href = `./unauthorized.html?page=${page}`;
}

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function formatDateTime(value) {
  if (!value) {
    return "-";
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return value;
  }

  return date.toLocaleString();
}

function setDetail(client) {
  if (!client) {
    detailFields.id.textContent = "-";
    detailFields.name.textContent = "Select a client";
    detailFields.postcode.textContent = "-";
    detailFields.email.textContent = "-";
    detailFields.phone.textContent = "-";
    detailFields.status.textContent = "-";
    detailFields.tags.textContent = "-";
    detailFields.visitCount.textContent = "-";
    detailFields.carerCount.textContent = "-";
    detailFields.lastVisitAt.textContent = "-";
    linkedCarersList.innerHTML = "";
    return;
  }

  detailFields.id.textContent = client.id || "-";
  detailFields.name.textContent = client.name || "-";
  detailFields.postcode.textContent = client.postcode || "-";
  detailFields.email.textContent = client.email || "-";
  detailFields.phone.textContent = client.phone || "-";
  detailFields.status.textContent = formatStatusLabel(client.status);
  detailFields.tags.textContent = formatTags(client.tags);
  detailFields.visitCount.textContent = String(client.relationships?.visitCount || 0);
  detailFields.carerCount.textContent = String(client.relationships?.carerCount || 0);
  detailFields.lastVisitAt.textContent = formatDateTime(client.relationships?.lastVisitAt);

  linkedCarersList.innerHTML = "";
  const carers = Array.isArray(client.relationships?.carers) ? client.relationships.carers : [];
  if (!carers.length) {
    const li = document.createElement("li");
    li.textContent = "No linked carers found.";
    linkedCarersList.appendChild(li);
    return;
  }

  for (const carer of carers) {
    const li = document.createElement("li");
    li.textContent = `${carer.name || "Unknown"} (${carer.id || "-"})`;
    linkedCarersList.appendChild(li);
  }
}

function normalizeStatus(value) {
  return String(value || "").trim().toLowerCase();
}

function collectStatusOptions(items) {
  const set = new Set();
  for (const item of items) {
    const status = normalizeStatus(item?.status);
    if (status) {
      set.add(status);
    }
  }
  return Array.from(set).sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
}

function formatStatusLabel(status) {
  const normalized = normalizeStatus(status);
  if (!normalized) {
    return "Unknown";
  }
  return normalized
    .split(/[\s_-]+/)
    .filter(Boolean)
    .map((part) => part[0].toUpperCase() + part.slice(1))
    .join(" ");
}

function formatTags(tags) {
  if (!Array.isArray(tags) || !tags.length) {
    return "-";
  }
  const unique = Array.from(
    new Set(
      tags
        .map((tag) => String(tag || "").trim())
        .filter(Boolean)
    )
  );
  return unique.length ? unique.join(", ") : "-";
}

function passesStatusFilter(status) {
  if (!selectedClientStatuses.size) {
    return true;
  }
  const normalized = normalizeStatus(status);
  if (!normalized) {
    return false;
  }
  return selectedClientStatuses.has(normalized);
}

function renderStatusFilters() {
  if (!clientStatusFilters) {
    return;
  }

  const options = collectStatusOptions(allClients);
  if (!options.length) {
    clientStatusFilters.hidden = true;
    return;
  }

  const nextSelected = new Set(Array.from(selectedClientStatuses).filter((status) => options.includes(status)));
  if (!nextSelected.size) {
    selectedClientStatuses = new Set(options.filter((status) => DEFAULT_STATUS_FILTERS.has(status)));
    if (!selectedClientStatuses.size) {
      selectedClientStatuses = new Set(options);
    }
  } else {
    selectedClientStatuses = nextSelected;
  }

  clientStatusFilters.hidden = false;
  clientStatusFilters.innerHTML = "";

  const allBtn = document.createElement("button");
  allBtn.type = "button";
  allBtn.className = `status-filter-btn${selectedClientStatuses.size === options.length ? " active" : ""}`;
  allBtn.textContent = "All";
  allBtn.addEventListener("click", () => {
    selectedClientStatuses = new Set(options);
    renderStatusFilters();
    renderClients();
  });
  clientStatusFilters.appendChild(allBtn);

  for (const status of options) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = `status-filter-btn${selectedClientStatuses.has(status) ? " active" : ""}`;
    btn.textContent = formatStatusLabel(status);
    btn.addEventListener("click", () => {
      const next = new Set(selectedClientStatuses);
      if (next.has(status)) {
        next.delete(status);
      } else {
        next.add(status);
      }
      if (!next.size) {
        next.add(status);
      }
      selectedClientStatuses = next;
      renderStatusFilters();
      renderClients();
    });
    clientStatusFilters.appendChild(btn);
  }
}

function getFilteredClients() {
  const query = String(searchInput.value || "").trim().toLowerCase();
  return allClients.filter((client) => {
    if (!passesStatusFilter(client.status)) {
      return false;
    }
    if (!query) {
      return true;
    }
    return (
      String(client.id || "").toLowerCase().includes(query) ||
      String(client.name || "").toLowerCase().includes(query) ||
      String(client.postcode || "").toLowerCase().includes(query) ||
      formatStatusLabel(client.status).toLowerCase().includes(query) ||
      formatTags(client.tags).toLowerCase().includes(query)
    );
  });
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function renderClients() {
  const filtered = getFilteredClients();
  clientsTableBody.innerHTML = "";

  if (!filtered.length) {
    emptyState.hidden = false;
    setDetail(null);
    return;
  }

  emptyState.hidden = true;

  const selected = filtered.find((client) => client.id === selectedClientId) || filtered[0];
  selectedClientId = selected.id;

  for (const client of filtered) {
    const tr = document.createElement("tr");
    tr.classList.toggle("selected", client.id === selectedClientId);

    tr.innerHTML = `
      <td>${escapeHtml(client.id)}</td>
      <td>${escapeHtml(client.name)}</td>
      <td>${escapeHtml(formatStatusLabel(client.status))}</td>
      <td>${escapeHtml(client.postcode || "-")}</td>
      <td>${escapeHtml(String(client.relationships?.carerCount || 0))}</td>
      <td>${escapeHtml(String(client.relationships?.visitCount || 0))}</td>
    `;

    tr.addEventListener("click", () => {
      selectedClientId = client.id;
      setDetail(client);
      renderClients();
    });

    clientsTableBody.appendChild(tr);
  }

  setDetail(selected);
}

async function init() {
  try {
    const restored = await authController.restoreSession();
    account = restored;
    if (!account) {
      window.location.href = "./index.html";
      return;
    }

    const profile = await directoryApi.getCurrentUser();
    const role = String(profile?.role || "").trim().toLowerCase();
    if (!canAccessPage(role, "clients")) {
      redirectToUnauthorized("clients");
      return;
    }
    renderTopNavigation({ role });

    setStatus("Loading clients...");
    const payload = await directoryApi.listOneTouchClients({ limit: 500 });
    allClients = Array.isArray(payload?.clients) ? payload.clients : [];

    const warnings = Array.isArray(payload?.warnings) ? payload.warnings.filter(Boolean) : [];
    warningState.hidden = warnings.length === 0;
    warningState.textContent = warnings.join(" ");

    renderStatusFilters();
    setStatus(`Loaded ${allClients.length} client(s).`);
    renderClients();
  } catch (error) {
    if (error?.status === 403) {
      redirectToUnauthorized("clients");
      return;
    }
    console.error(error);
    setStatus(error?.message || "Could not load clients.", true);
    emptyState.hidden = false;
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

searchInput?.addEventListener("input", () => {
  renderClients();
});

signOutBtn?.addEventListener("click", async () => {
  try {
    signOutBtn.disabled = true;
    await authController.signOut();
  } finally {
    window.location.href = "./index.html";
  }
});

void init();
