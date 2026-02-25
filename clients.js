import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";

const searchInput = document.getElementById("searchInput");
const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const clientsTableBody = document.getElementById("clientsTableBody");
const emptyState = document.getElementById("emptyState");
const detailRoot = document.getElementById("clientDetail");
const detailFields = {
  id: detailRoot?.querySelector('[data-field="id"]'),
  name: detailRoot?.querySelector('[data-field="name"]'),
  address: detailRoot?.querySelector('[data-field="address"]'),
  town: detailRoot?.querySelector('[data-field="town"]'),
  county: detailRoot?.querySelector('[data-field="county"]'),
  postcode: detailRoot?.querySelector('[data-field="postcode"]'),
  email: detailRoot?.querySelector('[data-field="email"]'),
};

const API_BASE_URL = (FRONTEND_CONFIG.apiBaseUrl || "").replace(/\/+$/, "");
const CLIENTS_ENDPOINT = API_BASE_URL ? `${API_BASE_URL}/api/clients` : "/api/clients";
const ME_ENDPOINT = API_BASE_URL ? `${API_BASE_URL}/api/auth/me` : "/api/auth/me";

let allClients = [];
let selectedClientId = "";
let selectedClient = null;
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

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function setDetail(client) {
  selectedClient = client || null;
  if (!client) {
    detailFields.id.textContent = "-";
    detailFields.name.textContent = "Select a client";
    detailFields.address.textContent = "-";
    detailFields.town.textContent = "-";
    detailFields.county.textContent = "-";
    detailFields.postcode.textContent = "-";
    detailFields.email.textContent = "-";
    return;
  }

  detailFields.id.textContent = client.id || "-";
  detailFields.name.textContent = client.name || "-";
  detailFields.address.textContent = client.address || "-";
  detailFields.town.textContent = client.town || "-";
  detailFields.county.textContent = client.county || "-";
  detailFields.postcode.textContent = client.postcode || "-";
  detailFields.email.textContent = client.email || "-";
}

function getFilteredClients() {
  const query = String(searchInput.value || "").trim().toLowerCase();
  if (!query) {
    return allClients;
  }

  return allClients.filter((client) => {
    return (
      String(client.id || "").toLowerCase().includes(query) ||
      String(client.name || "").toLowerCase().includes(query)
    );
  });
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

  for (const client of filtered) {
    const tr = document.createElement("tr");
    if (client.id === selectedClientId) {
      tr.classList.add("selected");
    }

    tr.innerHTML = `
      <td>${escapeHtml(client.id)}</td>
      <td>${escapeHtml(client.name)}</td>
    `;

    tr.addEventListener("click", () => {
      selectedClientId = client.id;
      setDetail(client);
      renderClients();
    });

    clientsTableBody.appendChild(tr);
  }

  const selected = filtered.find((client) => client.id === selectedClientId) || filtered[0];
  selectedClientId = selected.id;
  setDetail(selected);
  for (const row of clientsTableBody.querySelectorAll("tr")) {
    const idCell = row.children[0];
    row.classList.toggle("selected", idCell?.textContent === selected.id);
  }
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

async function fetchClients() {
  const token = await authController.acquireToken([FRONTEND_CONFIG.apiScope]);
  const response = await fetch(CLIENTS_ENDPOINT, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
    },
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Clients request failed (${response.status}): ${text || "Unknown error"}`);
  }

  const data = await response.json();
  return Array.isArray(data?.clients) ? data.clients : [];
}

async function fetchCurrentUser() {
  const token = await authController.acquireToken([FRONTEND_CONFIG.apiScope]);
  const response = await fetch(ME_ENDPOINT, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
    },
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Profile request failed (${response.status}): ${text || "Unknown error"}`);
  }

  return response.json();
}

async function loadClientsWithRetry() {
  const maxAttempts = 2;
  let lastError = null;

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      return await fetchClients();
    } catch (error) {
      lastError = error;
      if (attempt < maxAttempts) {
        await new Promise((resolve) => setTimeout(resolve, 450 * attempt));
      }
    }
  }

  throw lastError;
}

async function init() {
  try {
    const restored = await authController.restoreSession();
    account = restored;
    if (!account) {
      window.location.href = "./index.html";
      return;
    }

    const profile = await fetchCurrentUser();
    const role = String(profile?.role || "").trim().toLowerCase();
    if (role === "marketing") {
      window.location.href = "./marketing.html";
      return;
    }

    setStatus("Loading clients...");
    allClients = await loadClientsWithRetry();
    setStatus(`Loaded ${allClients.length} client(s).`);
    renderClients();
  } catch (error) {
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
