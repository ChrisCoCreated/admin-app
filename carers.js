import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { renderTopNavigation } from "./navigation.js?v=20260311";

const searchInput = document.getElementById("searchInput");
const areaFilterSelect = document.getElementById("areaFilterSelect");
const careCompFilterSelect = document.getElementById("careCompFilterSelect");
const statusFilterSelect = document.getElementById("statusFilterSelect");
const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const bulkTagSelect = document.getElementById("bulkTagSelect");
const selectedCountLabel = document.getElementById("selectedCountLabel");
const selectVisibleBtn = document.getElementById("selectVisibleBtn");
const clearSelectionBtn = document.getElementById("clearSelectionBtn");
const applyBulkTagBtn = document.getElementById("applyBulkTagBtn");
const bulkTagStatus = document.getElementById("bulkTagStatus");
const carersTableBody = document.getElementById("carersTableBody");
const emptyState = document.getElementById("emptyState");
const warningState = document.getElementById("warningState");
const detailRoot = document.getElementById("carerDetail");

const detailFields = {
  id: detailRoot?.querySelector('[data-field="id"]'),
  name: detailRoot?.querySelector('[data-field="name"]'),
  status: detailRoot?.querySelector('[data-field="status"]'),
  area: detailRoot?.querySelector('[data-field="area"]'),
  careComp: detailRoot?.querySelector('[data-field="careComp"]'),
  contractedHours: detailRoot?.querySelector('[data-field="contractedHours"]'),
  otherTags: detailRoot?.querySelector('[data-field="otherTags"]'),
  postcode: detailRoot?.querySelector('[data-field="postcode"]'),
  email: detailRoot?.querySelector('[data-field="email"]'),
  phone: detailRoot?.querySelector('[data-field="phone"]'),
};

let allCarers = [];
let selectedCarerId = "";
let selectedCarerIds = new Set();
let availableTags = [];
let bulkTagBusy = false;

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

function normalizeText(value) {
  return String(value || "").trim().toLowerCase();
}

function normalizeValue(value) {
  return String(value || "").trim();
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function setBulkTagStatus(message, isError = false) {
  if (!bulkTagStatus) {
    return;
  }
  bulkTagStatus.textContent = message;
  bulkTagStatus.classList.toggle("error", isError);
}

function getAreaLabel(carer) {
  return normalizeValue(carer.area) || "Unassigned";
}

function getStatusLabel(carer) {
  const normalized = normalizeValue(carer.status);
  if (!normalized) {
    return "Unknown";
  }

  return normalized
    .split(/[\s_-]+/)
    .filter(Boolean)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase())
    .join(" ");
}

function getCareCompLabel(carer) {
  return normalizeValue(carer.careCompanionshipTag) || "Unassigned";
}

function getOtherTagsLabel(carer) {
  const tags = Array.isArray(carer.otherTags) ? carer.otherTags.map(normalizeValue).filter(Boolean) : [];
  return tags.length ? tags.join(", ") : "-";
}

function formatHours(value) {
  const numeric = Number(value || 0);
  if (!Number.isFinite(numeric)) {
    return "0";
  }
  return numeric % 1 === 0 ? String(numeric) : numeric.toFixed(1);
}

function renderBulkTagOptions() {
  if (!bulkTagSelect) {
    return;
  }
  const currentValue = String(bulkTagSelect.value || "");
  bulkTagSelect.innerHTML = '<option value="">Select a tag</option>';
  for (const tag of availableTags) {
    const option = document.createElement("option");
    option.value = String(tag.id);
    option.textContent = tag.name || `Tag ${tag.id}`;
    bulkTagSelect.appendChild(option);
  }
  bulkTagSelect.value = availableTags.some((tag) => String(tag.id) === currentValue) ? currentValue : "";
}

function updateBulkTagControls() {
  const selectedCount = selectedCarerIds.size;
  if (selectedCountLabel) {
    selectedCountLabel.textContent = `${selectedCount} selected`;
  }
  if (clearSelectionBtn) {
    clearSelectionBtn.disabled = selectedCount === 0 || bulkTagBusy;
  }
  if (selectVisibleBtn) {
    selectVisibleBtn.disabled = getFilteredCarers().length === 0 || bulkTagBusy;
  }
  if (applyBulkTagBtn) {
    applyBulkTagBtn.disabled = bulkTagBusy || selectedCount === 0 || !String(bulkTagSelect?.value || "").trim();
  }
}

function setBulkTagBusy(isBusy) {
  bulkTagBusy = Boolean(isBusy);
  if (bulkTagSelect) {
    bulkTagSelect.disabled = bulkTagBusy;
  }
  updateBulkTagControls();
}

function setDetail(carer) {
  if (!carer) {
    detailFields.id.textContent = "-";
    detailFields.name.textContent = "Select a carer";
    detailFields.status.textContent = "-";
    detailFields.area.textContent = "-";
    detailFields.careComp.textContent = "-";
    detailFields.contractedHours.textContent = "-";
    detailFields.otherTags.textContent = "-";
    detailFields.postcode.textContent = "-";
    detailFields.email.textContent = "-";
    detailFields.phone.textContent = "-";
    return;
  }

  detailFields.id.textContent = normalizeValue(carer.id) || "-";
  detailFields.name.textContent = normalizeValue(carer.name) || "-";
  detailFields.status.textContent = getStatusLabel(carer);
  detailFields.area.textContent = getAreaLabel(carer);
  detailFields.careComp.textContent = getCareCompLabel(carer);
  detailFields.contractedHours.textContent = formatHours(carer.contractedHours);
  detailFields.otherTags.textContent = getOtherTagsLabel(carer);
  detailFields.postcode.textContent = normalizeValue(carer.postcode) || "-";
  detailFields.email.textContent = normalizeValue(carer.email) || "-";
  detailFields.phone.textContent = normalizeValue(carer.phone) || "-";
}

function renderFilterOptions() {
  const areaOptions = Array.from(new Set(allCarers.map(getAreaLabel))).sort((a, b) => a.localeCompare(b));
  const careCompOptions = Array.from(new Set(allCarers.map(getCareCompLabel))).sort((a, b) => a.localeCompare(b));
  const statusOptions = Array.from(new Set(allCarers.map(getStatusLabel))).sort((a, b) => a.localeCompare(b));

  const currentArea = String(areaFilterSelect.value || "all");
  const currentCare = String(careCompFilterSelect.value || "all");
  const currentStatus = String(statusFilterSelect?.value || "all");

  areaFilterSelect.innerHTML = '<option value="all">All areas</option>';
  for (const area of areaOptions) {
    const option = document.createElement("option");
    option.value = area;
    option.textContent = area;
    areaFilterSelect.appendChild(option);
  }

  careCompFilterSelect.innerHTML = '<option value="all">All tags</option>';
  for (const tag of careCompOptions) {
    const option = document.createElement("option");
    option.value = tag;
    option.textContent = tag;
    careCompFilterSelect.appendChild(option);
  }

  if (statusFilterSelect) {
    statusFilterSelect.innerHTML = '<option value="all">All statuses</option>';
    for (const status of statusOptions) {
      const option = document.createElement("option");
      option.value = status;
      option.textContent = status;
      statusFilterSelect.appendChild(option);
    }
    statusFilterSelect.value = statusOptions.includes(currentStatus) ? currentStatus : "all";
  }

  areaFilterSelect.value = areaOptions.includes(currentArea) ? currentArea : "all";
  careCompFilterSelect.value = careCompOptions.includes(currentCare) ? currentCare : "all";
}

function getFilteredCarers() {
  const query = normalizeText(searchInput.value);
  const selectedArea = String(areaFilterSelect.value || "all");
  const selectedCareComp = String(careCompFilterSelect.value || "all");
  const selectedStatus = String(statusFilterSelect?.value || "all");

  return allCarers.filter((carer) => {
    if (selectedArea !== "all" && getAreaLabel(carer) !== selectedArea) {
      return false;
    }
    if (selectedCareComp !== "all" && getCareCompLabel(carer) !== selectedCareComp) {
      return false;
    }
    if (selectedStatus !== "all" && getStatusLabel(carer) !== selectedStatus) {
      return false;
    }

    if (!query) {
      return true;
    }

    return (
      normalizeText(carer.id).includes(query) ||
      normalizeText(carer.name).includes(query) ||
      normalizeText(carer.postcode).includes(query) ||
      normalizeText(getStatusLabel(carer)).includes(query) ||
      normalizeText(getAreaLabel(carer)).includes(query) ||
      normalizeText(getCareCompLabel(carer)).includes(query) ||
      normalizeText(getOtherTagsLabel(carer)).includes(query)
    );
  });
}

function renderCarers() {
  const filtered = getFilteredCarers();
  carersTableBody.innerHTML = "";

  if (!filtered.length) {
    emptyState.hidden = false;
    setDetail(null);
    updateBulkTagControls();
    return;
  }

  emptyState.hidden = true;

  const selected = filtered.find((carer) => carer.id === selectedCarerId) || filtered[0];
  selectedCarerId = selected.id;

  for (const carer of filtered) {
    const tr = document.createElement("tr");
    tr.classList.toggle("selected", carer.id === selectedCarerId);
    tr.innerHTML = `
      <td class="selection-cell"></td>
      <td>${escapeHtml(carer.id)}</td>
      <td>${escapeHtml(carer.name)}</td>
      <td>${escapeHtml(getStatusLabel(carer))}</td>
      <td>${escapeHtml(getAreaLabel(carer))}</td>
      <td>${escapeHtml(getCareCompLabel(carer))}</td>
      <td>${escapeHtml(formatHours(carer.contractedHours))}</td>
    `;

    const selectionCell = tr.querySelector(".selection-cell");
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = selectedCarerIds.has(carer.id);
    checkbox.disabled = bulkTagBusy;
    checkbox.setAttribute("aria-label", `Select ${carer.name || carer.id || "carer"}`);
    checkbox.addEventListener("click", (event) => {
      event.stopPropagation();
    });
    checkbox.addEventListener("change", () => {
      if (checkbox.checked) {
        selectedCarerIds.add(carer.id);
      } else {
        selectedCarerIds.delete(carer.id);
      }
      updateBulkTagControls();
    });
    selectionCell?.appendChild(checkbox);

    tr.addEventListener("click", () => {
      selectedCarerId = carer.id;
      setDetail(carer);
      renderCarers();
    });

    carersTableBody.appendChild(tr);
  }

  setDetail(selected);
  updateBulkTagControls();
}

async function refreshCarersData() {
  setStatus("Loading carers...");
  const payload = await directoryApi.listCarers({ limit: 500 });
  allCarers = Array.isArray(payload?.carers) ? payload.carers : [];
  const warnings = Array.isArray(payload?.warnings) ? payload.warnings.filter(Boolean) : [];
  warningState.hidden = warnings.length === 0;
  warningState.textContent = warnings.join(" ");
  renderFilterOptions();
  renderCarers();
  setStatus(`Loaded ${allCarers.length} carer(s).`);
}

async function loadTagCatalog() {
  const payload = await directoryApi.getOneTouchTags();
  availableTags = Array.isArray(payload?.tags) ? payload.tags : [];
  renderBulkTagOptions();
  updateBulkTagControls();
}

async function init() {
  try {
    const account = await authController.restoreSession();
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
    renderTopNavigation({ role });

    await Promise.all([refreshCarersData(), loadTagCatalog()]);
  } catch (error) {
    console.error(error);
    setStatus(error?.message || "Could not load carers.", true);
    emptyState.hidden = false;
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

searchInput?.addEventListener("input", renderCarers);
areaFilterSelect?.addEventListener("change", renderCarers);
careCompFilterSelect?.addEventListener("change", renderCarers);
statusFilterSelect?.addEventListener("change", renderCarers);
bulkTagSelect?.addEventListener("change", updateBulkTagControls);

selectVisibleBtn?.addEventListener("click", () => {
  for (const carer of getFilteredCarers()) {
    selectedCarerIds.add(carer.id);
  }
  setBulkTagStatus(`Selected ${selectedCarerIds.size} carer(s).`);
  renderCarers();
  updateBulkTagControls();
});

clearSelectionBtn?.addEventListener("click", () => {
  selectedCarerIds = new Set();
  setBulkTagStatus("Selection cleared.");
  renderCarers();
  updateBulkTagControls();
});

applyBulkTagBtn?.addEventListener("click", async () => {
  const recordIds = Array.from(selectedCarerIds);
  const tagId = Number(bulkTagSelect?.value || 0);
  if (!recordIds.length || !Number.isFinite(tagId) || tagId <= 0) {
    updateBulkTagControls();
    return;
  }

  setBulkTagBusy(true);
  setBulkTagStatus(`Applying tag to ${recordIds.length} carer(s)...`);
  try {
    const result = await directoryApi.applyBulkCarerTag({ recordIds, tagId });
    await refreshCarersData();
    const failed = Number(result?.failed || 0);
    const succeeded = Number(result?.succeeded || 0);
    if (failed > 0) {
      selectedCarerIds = new Set(
        Array.isArray(result?.results) ? result.results.filter((item) => !item?.ok).map((item) => item?.id).filter(Boolean) : []
      );
      const failedNames = Array.isArray(result?.results)
        ? result.results
            .filter((item) => !item?.ok)
            .slice(0, 3)
            .map((item) => item?.name || item?.id || "Unknown")
            .join(", ")
        : "";
      setBulkTagStatus(
        `Bulk tag complete: ${succeeded} succeeded, ${failed} failed.${failedNames ? ` Failed: ${failedNames}` : ""}`,
        true
      );
    } else {
      selectedCarerIds = new Set();
      setBulkTagStatus(`Bulk tag complete: ${succeeded} carer(s) updated.`);
    }
    if (bulkTagSelect) {
      bulkTagSelect.value = "";
    }
    renderCarers();
  } catch (error) {
    console.error(error);
    setBulkTagStatus(error?.message || "Could not apply bulk tag.", true);
  } finally {
    setBulkTagBusy(false);
  }
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
