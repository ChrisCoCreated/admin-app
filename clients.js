import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js";

const searchInput = document.getElementById("searchInput");
const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
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
  visitCount: detailRoot?.querySelector('[data-field="visitCount"]'),
  carerCount: detailRoot?.querySelector('[data-field="carerCount"]'),
  lastVisitAt: detailRoot?.querySelector('[data-field="lastVisitAt"]'),
};

let allClients = [];
let selectedClientId = "";
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

function getFilteredClients() {
  const query = String(searchInput.value || "").trim().toLowerCase();
  if (!query) {
    return allClients;
  }

  return allClients.filter((client) => {
    return (
      String(client.id || "").toLowerCase().includes(query) ||
      String(client.name || "").toLowerCase().includes(query) ||
      String(client.postcode || "").toLowerCase().includes(query)
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
