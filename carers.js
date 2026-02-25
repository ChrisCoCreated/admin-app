import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";

const searchInput = document.getElementById("searchInput");
const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const carersTableBody = document.getElementById("carersTableBody");
const emptyState = document.getElementById("emptyState");
const warningState = document.getElementById("warningState");
const detailRoot = document.getElementById("carerDetail");
const linkedClientsList = document.getElementById("linkedClientsList");

const detailFields = {
  id: detailRoot?.querySelector('[data-field="id"]'),
  name: detailRoot?.querySelector('[data-field="name"]'),
  postcode: detailRoot?.querySelector('[data-field="postcode"]'),
  email: detailRoot?.querySelector('[data-field="email"]'),
  phone: detailRoot?.querySelector('[data-field="phone"]'),
  visitCount: detailRoot?.querySelector('[data-field="visitCount"]'),
  clientCount: detailRoot?.querySelector('[data-field="clientCount"]'),
  lastVisitAt: detailRoot?.querySelector('[data-field="lastVisitAt"]'),
};

let allCarers = [];
let selectedCarerId = "";
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

function setDetail(carer) {
  if (!carer) {
    detailFields.id.textContent = "-";
    detailFields.name.textContent = "Select a carer";
    detailFields.postcode.textContent = "-";
    detailFields.email.textContent = "-";
    detailFields.phone.textContent = "-";
    detailFields.visitCount.textContent = "-";
    detailFields.clientCount.textContent = "-";
    detailFields.lastVisitAt.textContent = "-";
    linkedClientsList.innerHTML = "";
    return;
  }

  detailFields.id.textContent = carer.id || "-";
  detailFields.name.textContent = carer.name || "-";
  detailFields.postcode.textContent = carer.postcode || "-";
  detailFields.email.textContent = carer.email || "-";
  detailFields.phone.textContent = carer.phone || "-";
  detailFields.visitCount.textContent = String(carer.relationships?.visitCount || 0);
  detailFields.clientCount.textContent = String(carer.relationships?.clientCount || 0);
  detailFields.lastVisitAt.textContent = formatDateTime(carer.relationships?.lastVisitAt);

  linkedClientsList.innerHTML = "";
  const clients = Array.isArray(carer.relationships?.clients) ? carer.relationships.clients : [];
  if (!clients.length) {
    const li = document.createElement("li");
    li.textContent = "No linked clients found.";
    linkedClientsList.appendChild(li);
    return;
  }

  for (const client of clients) {
    const li = document.createElement("li");
    li.textContent = `${client.name || "Unknown"} (${client.id || "-"})`;
    linkedClientsList.appendChild(li);
  }
}

function getFilteredCarers() {
  const query = String(searchInput.value || "").trim().toLowerCase();
  if (!query) {
    return allCarers;
  }

  return allCarers.filter((carer) => {
    return (
      String(carer.id || "").toLowerCase().includes(query) ||
      String(carer.name || "").toLowerCase().includes(query) ||
      String(carer.postcode || "").toLowerCase().includes(query)
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

function renderCarers() {
  const filtered = getFilteredCarers();
  carersTableBody.innerHTML = "";

  if (!filtered.length) {
    emptyState.hidden = false;
    setDetail(null);
    return;
  }

  emptyState.hidden = true;

  const selected = filtered.find((carer) => carer.id === selectedCarerId) || filtered[0];
  selectedCarerId = selected.id;

  for (const carer of filtered) {
    const tr = document.createElement("tr");
    tr.classList.toggle("selected", carer.id === selectedCarerId);

    tr.innerHTML = `
      <td>${escapeHtml(carer.id)}</td>
      <td>${escapeHtml(carer.name)}</td>
      <td>${escapeHtml(carer.postcode || "-")}</td>
      <td>${escapeHtml(String(carer.relationships?.clientCount || 0))}</td>
      <td>${escapeHtml(String(carer.relationships?.visitCount || 0))}</td>
    `;

    tr.addEventListener("click", () => {
      selectedCarerId = carer.id;
      setDetail(carer);
      renderCarers();
    });

    carersTableBody.appendChild(tr);
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
    if (role === "marketing") {
      window.location.href = "./marketing.html";
      return;
    }

    setStatus("Loading carers...");
    const payload = await directoryApi.listCarers({ limit: 500 });
    allCarers = Array.isArray(payload?.carers) ? payload.carers : [];

    const warnings = Array.isArray(payload?.warnings) ? payload.warnings.filter(Boolean) : [];
    warningState.hidden = warnings.length === 0;
    warningState.textContent = warnings.join(" ");

    setStatus(`Loaded ${allCarers.length} carer(s).`);
    renderCarers();
  } catch (error) {
    console.error(error);
    setStatus(error?.message || "Could not load carers.", true);
    emptyState.hidden = false;
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

searchInput?.addEventListener("input", () => {
  renderCarers();
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
