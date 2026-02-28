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
const reconcilePanel = document.getElementById("reconcilePanel");
const reconcileRefreshBtn = document.getElementById("reconcileRefreshBtn");
const reconcileStatus = document.getElementById("reconcileStatus");
const copyIdBody = document.getElementById("copyIdBody");
const missingBody = document.getElementById("missingBody");
const updateBody = document.getElementById("updateBody");
const ambiguousBody = document.getElementById("ambiguousBody");
const errorsBody = document.getElementById("errorsBody");
const copyAllBtn = document.getElementById("copyAllBtn");
const updateAllBtn = document.getElementById("updateAllBtn");

const detailFields = {
  id: detailRoot?.querySelector('[data-field="id"]'),
  name: detailRoot?.querySelector('[data-field="name"]'),
  postcode: detailRoot?.querySelector('[data-field="postcode"]'),
  email: detailRoot?.querySelector('[data-field="email"]'),
  phone: detailRoot?.querySelector('[data-field="phone"]'),
  status: detailRoot?.querySelector('[data-field="status"]'),
  tags: detailRoot?.querySelector('[data-field="tags"]'),
};

const DEFAULT_STATUS_FILTERS = new Set(["active", "pending"]);
const STATUS_FILTER_ORDER = ["active", "pending", "archived"];

let allClients = [];
let selectedClientId = "";
let selectedClientStatuses = new Set(DEFAULT_STATUS_FILTERS);
let account = null;
let currentRole = "";
let reconcilePreview = null;
let reconcileBusy = false;

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

function setDetail(client) {
  if (!client) {
    detailFields.id.textContent = "-";
    detailFields.name.textContent = "Select a client";
    detailFields.postcode.textContent = "-";
    detailFields.email.textContent = "-";
    detailFields.phone.textContent = "-";
    detailFields.status.textContent = "-";
    detailFields.tags.textContent = "-";
    return;
  }

  detailFields.id.textContent = client.id || "-";
  detailFields.name.textContent = client.name || "-";
  detailFields.postcode.textContent = client.postcode || "-";
  detailFields.email.textContent = client.email || "-";
  detailFields.phone.textContent = client.phone || "-";
  detailFields.status.textContent = formatStatusLabel(client.status);
  detailFields.tags.textContent = formatTags(client.tags);
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
  if (!set.has("active")) {
    set.add("active");
  }
  if (!set.has("pending")) {
    set.add("pending");
  }
  if (!set.has("archived")) {
    set.add("archived");
  }

  const preferred = STATUS_FILTER_ORDER.filter((status) => set.has(status));
  const remaining = Array.from(set)
    .filter((status) => !preferred.includes(status))
    .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
  return [...preferred, ...remaining];
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

function formatBooleanIndicator(value) {
  if (value === true) {
    return { klass: "is-yes", label: "Yes" };
  }
  if (value === false) {
    return { klass: "is-no", label: "No" };
  }
  return { klass: "is-unknown", label: "Unknown" };
}

function renderIndicator(value, fieldLabel) {
  const indicator = formatBooleanIndicator(value);
  return `<span class="client-indicator ${indicator.klass}" role="img" aria-label="${escapeHtml(
    `${fieldLabel}: ${indicator.label}`
  )}" title="${escapeHtml(indicator.label)}"></span>`;
}

function renderXeroLink(xeroId) {
  const id = String(xeroId || "").trim();
  if (!id) {
    return "-";
  }
  const url = `https://go.xero.com/app/!q4T5z/contacts/contact/${encodeURIComponent(id)}`;
  return `<a class="xero-link-btn" href="${escapeHtml(url)}" target="_blank" rel="noopener noreferrer">Xero</a>`;
}

function canManageReconciliation() {
  return currentRole === "admin" || currentRole === "care_manager";
}

function setReconcileStatus(message, isError = false) {
  if (!reconcileStatus) {
    return;
  }
  reconcileStatus.textContent = message;
  reconcileStatus.classList.toggle("error", isError);
}

function setReconcileBusy(isBusy) {
  reconcileBusy = isBusy;
  if (reconcileRefreshBtn) {
    reconcileRefreshBtn.disabled = isBusy;
  }
  if (copyAllBtn) {
    const hasItems =
      Array.isArray(reconcilePreview?.copyOneTouchIdCandidates) && reconcilePreview.copyOneTouchIdCandidates.length > 0;
    copyAllBtn.disabled = isBusy || !hasItems;
  }
  if (updateAllBtn) {
    const hasItems = Array.isArray(reconcilePreview?.updateCandidates) && reconcilePreview.updateCandidates.length > 0;
    updateAllBtn.disabled = isBusy || !hasItems;
  }
  if (reconcilePanel) {
    const controls = reconcilePanel.querySelectorAll(".recon-action-btn, .recon-select");
    for (const control of controls) {
      if (control.classList.contains("recon-action-btn")) {
        const baseDisabled = control.dataset.baseDisabled === "1";
        control.disabled = isBusy || baseDisabled;
      } else {
        control.disabled = isBusy;
      }
    }
  }
}

function renderEmptyRow(root, message, colspan = 4) {
  if (!root) {
    return;
  }
  root.innerHTML = "";
  const tr = document.createElement("tr");
  tr.innerHTML = `<td colspan="${colspan}" class="muted">${escapeHtml(message)}</td>`;
  root.appendChild(tr);
}

function createActionButton(label, onClick, disabled = false) {
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "secondary recon-action-btn";
  btn.textContent = label;
  btn.dataset.baseDisabled = disabled ? "1" : "0";
  btn.disabled = disabled;
  btn.addEventListener("click", async (event) => {
    event.preventDefault();
    await onClick();
  });
  return btn;
}

async function refreshClientsData() {
  setStatus("Loading clients...");
  const payload = await directoryApi.listOneTouchClients({ limit: 500 });
  allClients = Array.isArray(payload?.clients) ? payload.clients : [];
  const warnings = Array.isArray(payload?.warnings) ? payload.warnings.filter(Boolean) : [];
  warningState.hidden = warnings.length === 0;
  warningState.textContent = warnings.join(" ");
  renderStatusFilters();
  setStatus(`Loaded ${allClients.length} client(s).`);
  renderClients();
}

async function loadReconcilePreview() {
  if (!canManageReconciliation() || !reconcilePanel) {
    return;
  }
  setReconcileBusy(true);
  setReconcileStatus("Loading reconciliation preview...");
  try {
    reconcilePreview = await directoryApi.getClientsReconcilePreview();
    renderReconcilePreview();
    const totals = reconcilePreview?.totals || {};
    setReconcileStatus(
      `Copy: ${totals.copyOneTouchIdCandidates || 0}, Missing: ${totals.missingInSharePoint || 0}, ` +
        `Update: ${totals.updateCandidates || 0}, Ambiguous: ${totals.ambiguousMatches || 0}.`
    );
  } catch (error) {
    console.error(error);
    setReconcileStatus(error?.message || "Could not load reconciliation preview.", true);
    renderEmptyRow(copyIdBody, "Could not load data.");
    renderEmptyRow(missingBody, "Could not load data.");
    renderEmptyRow(updateBody, "Could not load data.");
    renderEmptyRow(ambiguousBody, "Could not load data.");
    renderEmptyRow(errorsBody, "Could not load data.", 3);
  } finally {
    setReconcileBusy(false);
  }
}

async function applyReconcileAction(payload, successMessage) {
  if (reconcileBusy) {
    return;
  }
  setReconcileBusy(true);
  setReconcileStatus("Applying reconciliation change...");
  try {
    await directoryApi.applyClientsReconcileAction(payload);
    setReconcileStatus(successMessage || "Reconciliation action applied.");
    await Promise.all([refreshClientsData(), loadReconcilePreview()]);
  } catch (error) {
    console.error(error);
    setReconcileStatus(error?.message || "Could not apply reconciliation action.", true);
  } finally {
    setReconcileBusy(false);
  }
}

function renderCopyCandidates(items) {
  if (!copyIdBody) {
    return;
  }
  copyIdBody.innerHTML = "";
  if (!items.length) {
    renderEmptyRow(copyIdBody, "No copy candidates.");
    if (copyAllBtn) {
      copyAllBtn.disabled = true;
    }
    return;
  }

  for (const item of items) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(item?.sharePoint?.name || "-")}</td>
      <td>${escapeHtml(item?.sharePoint?.dateOfBirth || "-")}</td>
      <td>${escapeHtml(item?.oneTouch?.name || "-")} (${escapeHtml(item?.oneTouch?.id || "-")})${
        item?.matchType === "fuzzy" ? " [fuzzy]" : ""
      }</td>
      <td></td>
    `;
    const actionCell = tr.lastElementChild;
    actionCell.appendChild(
      createActionButton("Copy", () =>
        applyReconcileAction(
          {
            action: "copy_onetouch_id",
            sharePointItemId: item?.sharePoint?.itemId || "",
            oneTouchClientId: item?.oneTouch?.id || "",
            expectedFingerprint: item?.expectedFingerprint || "",
          },
          "OneTouchID copied."
        )
      )
    );
    copyIdBody.appendChild(tr);
  }

  if (copyAllBtn) {
    copyAllBtn.disabled = reconcileBusy || items.length === 0;
  }
}

function renderMissingCandidates(items) {
  if (!missingBody) {
    return;
  }
  missingBody.innerHTML = "";
  if (!items.length) {
    renderEmptyRow(missingBody, "No missing records.");
    return;
  }

  for (const item of items) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(item?.oneTouch?.id || "-")}</td>
      <td>${escapeHtml(item?.oneTouch?.name || "-")}</td>
      <td>${escapeHtml(item?.oneTouch?.dateOfBirth || "-")}</td>
      <td></td>
    `;
    const actionCell = tr.lastElementChild;
    actionCell.appendChild(
      createActionButton("Add", () =>
        applyReconcileAction(
          {
            action: "add_missing",
            oneTouchClientId: item?.oneTouch?.id || "",
            expectedFingerprint: item?.expectedFingerprint || "",
          },
          "Missing SharePoint record added."
        )
      )
    );
    missingBody.appendChild(tr);
  }
}

function renderUpdateCandidates(items) {
  if (!updateBody) {
    return;
  }
  updateBody.innerHTML = "";
  if (!items.length) {
    renderEmptyRow(updateBody, "No update candidates.");
    if (updateAllBtn) {
      updateAllBtn.disabled = true;
    }
    return;
  }

  for (const item of items) {
    const diffs = Array.isArray(item?.differences) ? item.differences : [];
    const fieldsLabel = diffs.map((diff) => diff.field).join(", ") || "-";
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(item?.sharePoint?.name || "-")}</td>
      <td>${escapeHtml(item?.oneTouch?.id || "-")}</td>
      <td>${escapeHtml(fieldsLabel)}</td>
      <td></td>
    `;
    const actionCell = tr.lastElementChild;
    actionCell.appendChild(
      createActionButton("Update", () =>
        applyReconcileAction(
          {
            action: "update_record",
            sharePointItemId: item?.sharePoint?.itemId || "",
            oneTouchClientId: item?.oneTouch?.id || "",
            expectedFingerprint: item?.expectedFingerprint || "",
          },
          "SharePoint record updated."
        )
      )
    );
    updateBody.appendChild(tr);
  }

  if (updateAllBtn) {
    updateAllBtn.disabled = reconcileBusy || items.length === 0;
  }
}

function renderAmbiguousCandidates(items) {
  if (!ambiguousBody) {
    return;
  }
  ambiguousBody.innerHTML = "";
  if (!items.length) {
    renderEmptyRow(ambiguousBody, "No ambiguous matches.");
    return;
  }

  for (const item of items) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(item?.sharePoint?.name || "-")}</td>
      <td>${escapeHtml(item?.sharePoint?.dateOfBirth || "-")}</td>
      <td></td>
      <td></td>
    `;

    const selectCell = tr.children[2];
    const actionCell = tr.children[3];

    const select = document.createElement("select");
    select.className = "recon-select";
    const candidates = Array.isArray(item?.oneTouchCandidates) ? item.oneTouchCandidates : [];
    for (const candidate of candidates) {
      const option = document.createElement("option");
      option.value = String(candidate?.id || "");
      option.textContent = `${candidate?.name || "Unknown"} (${candidate?.id || "-"})`;
      select.appendChild(option);
    }
    selectCell.appendChild(select);

    actionCell.appendChild(
      createActionButton("Copy", () =>
        applyReconcileAction(
          {
            action: "copy_onetouch_id",
            sharePointItemId: item?.sharePoint?.itemId || "",
            oneTouchClientId: select.value,
            expectedFingerprint: item?.expectedFingerprint || "",
          },
          "OneTouchID copied."
        )
      )
    );

    ambiguousBody.appendChild(tr);
  }
}

function renderReconcilePreview() {
  const payload = reconcilePreview || {};
  renderCopyCandidates(Array.isArray(payload.copyOneTouchIdCandidates) ? payload.copyOneTouchIdCandidates : []);
  renderMissingCandidates(Array.isArray(payload.missingInSharePoint) ? payload.missingInSharePoint : []);
  renderUpdateCandidates(Array.isArray(payload.updateCandidates) ? payload.updateCandidates : []);
  renderAmbiguousCandidates(Array.isArray(payload.ambiguousMatches) ? payload.ambiguousMatches : []);
  renderErrors(Array.isArray(payload.errors) ? payload.errors : []);
}

function renderErrors(items) {
  if (!errorsBody) {
    return;
  }
  errorsBody.innerHTML = "";
  if (!items.length) {
    renderEmptyRow(errorsBody, "No unmatched SharePoint records.", 3);
    return;
  }

  for (const item of items) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(item?.sharePoint?.name || "-")}</td>
      <td>${escapeHtml(item?.sharePoint?.oneTouchId || "-")}</td>
      <td>${escapeHtml(item?.message || item?.type || "No matching OneTouch record.")}</td>
    `;
    errorsBody.appendChild(tr);
  }
}

async function applyUpdateAllCandidates() {
  const items = Array.isArray(reconcilePreview?.updateCandidates) ? reconcilePreview.updateCandidates : [];
  if (!items.length || reconcileBusy) {
    return;
  }

  setReconcileBusy(true);
  setReconcileStatus(`Updating ${items.length} record(s)...`);
  let updated = 0;
  let failed = 0;
  let lastError = "";

  try {
    for (const item of items) {
      try {
        const response = await directoryApi.applyClientsReconcileAction({
          action: "update_record",
          sharePointItemId: item?.sharePoint?.itemId || "",
          oneTouchClientId: item?.oneTouch?.id || "",
          expectedFingerprint: item?.expectedFingerprint || "",
        });
        if (response?.result?.updated === false) {
          continue;
        }
        updated += 1;
      } catch (error) {
        failed += 1;
        lastError = error?.message || String(error);
      }
    }
  } finally {
    setReconcileBusy(false);
  }

  await Promise.all([refreshClientsData(), loadReconcilePreview()]);
  if (failed > 0) {
    setReconcileStatus(
      `Update all complete: ${updated} updated, ${failed} failed.${lastError ? ` Last error: ${lastError}` : ""}`,
      true
    );
    return;
  }
  setReconcileStatus(`Update all complete: ${updated} updated.`);
}

async function applyCopyAllCandidates() {
  const items = Array.isArray(reconcilePreview?.copyOneTouchIdCandidates) ? reconcilePreview.copyOneTouchIdCandidates : [];
  if (!items.length || reconcileBusy) {
    return;
  }

  setReconcileBusy(true);
  setReconcileStatus(`Copying OneTouchID for ${items.length} record(s)...`);
  let copied = 0;
  let failed = 0;
  let lastError = "";

  try {
    for (const item of items) {
      try {
        const response = await directoryApi.applyClientsReconcileAction({
          action: "copy_onetouch_id",
          sharePointItemId: item?.sharePoint?.itemId || "",
          oneTouchClientId: item?.oneTouch?.id || "",
          expectedFingerprint: item?.expectedFingerprint || "",
        });
        if (response?.result?.updated === false) {
          continue;
        }
        copied += 1;
      } catch (error) {
        failed += 1;
        lastError = error?.message || String(error);
      }
    }
  } finally {
    setReconcileBusy(false);
  }

  await Promise.all([refreshClientsData(), loadReconcilePreview()]);
  if (failed > 0) {
    setReconcileStatus(
      `Copy all complete: ${copied} copied, ${failed} failed.${lastError ? ` Last error: ${lastError}` : ""}`,
      true
    );
    return;
  }
  setReconcileStatus(`Copy all complete: ${copied} copied.`);
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
      <td>${renderXeroLink(client.xeroId)}</td>
      <td class="indicator-cell">${renderIndicator(client.hasMandate, "Mandate")}</td>
      <td class="indicator-cell">${renderIndicator(client.hasMarketingConsent, "Marketing Consent")}</td>
    `;

    const xeroLink = tr.querySelector(".xero-link-btn");
    if (xeroLink) {
      xeroLink.addEventListener("click", (event) => {
        event.stopPropagation();
      });
    }

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
    currentRole = role;
    if (!canAccessPage(role, "clients")) {
      redirectToUnauthorized("clients");
      return;
    }
    renderTopNavigation({ role });

    await refreshClientsData();

    if (reconcilePanel) {
      reconcilePanel.hidden = !canManageReconciliation();
      if (canManageReconciliation()) {
        await loadReconcilePreview();
      }
    }
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

reconcileRefreshBtn?.addEventListener("click", async () => {
  await loadReconcilePreview();
});

copyAllBtn?.addEventListener("click", async () => {
  await applyCopyAllCandidates();
});

updateAllBtn?.addEventListener("click", async () => {
  await applyUpdateAllCandidates();
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
