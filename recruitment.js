import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260311";

const searchInput = document.getElementById("searchInput");
const locationFilterSelect = document.getElementById("locationFilterSelect");
const statusFilterSelect = document.getElementById("statusFilterSelect");
const sourceFilterSelect = document.getElementById("sourceFilterSelect");
const recruitmentTableBody = document.getElementById("recruitmentTableBody");
const emptyState = document.getElementById("emptyState");
const statusMessage = document.getElementById("statusMessage");
const signOutBtn = document.getElementById("signOutBtn");
const detailRoot = document.getElementById("candidateDetail");
const sharePointListLink = document.getElementById("sharePointListLink");
const importDropZone = document.getElementById("importDropZone");
const importFileInput = document.getElementById("importFileInput");
const importFileName = document.getElementById("importFileName");
const importSummary = document.getElementById("importSummary");
const importErrors = document.getElementById("importErrors");
const runImportBtn = document.getElementById("runImportBtn");
const importPreviewWrap = document.getElementById("importPreviewWrap");
const importPreviewTitle = document.getElementById("importPreviewTitle");
const importPreviewBody = document.getElementById("importPreviewBody");

const detailFields = {
  candidateName: detailRoot?.querySelector('[data-field="candidateName"]'),
  location: detailRoot?.querySelector('[data-field="location"]'),
  status: detailRoot?.querySelector('[data-field="status"]'),
  source: detailRoot?.querySelector('[data-field="source"]'),
  phoneNumber: detailRoot?.querySelector('[data-field="phoneNumber"]'),
  interviewBooked: detailRoot?.querySelector('[data-field="interviewBooked"]'),
  interviewWith: detailRoot?.querySelector('[data-field="interviewWith"]'),
  keepInMind: detailRoot?.querySelector('[data-field="keepInMind"]'),
  livesIn: detailRoot?.querySelector('[data-field="livesIn"]'),
  firstInterviewDate: detailRoot?.querySelector('[data-field="firstInterviewDate"]'),
  earmarkedFor: detailRoot?.querySelector('[data-field="earmarkedFor"]'),
  created: detailRoot?.querySelector('[data-field="created"]'),
  oneTouchLink: detailRoot?.querySelector('[data-field="oneTouchLink"]'),
  notes: detailRoot?.querySelector('[data-field="notes"]'),
};

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

let allCandidates = [];
let selectedCandidateId = "";
let addToOneTouchBusy = false;
let importBusy = false;
let pendingImportRows = [];
const IMPORT_PREVIEW_LIMIT = 10;

function normalizeText(value) {
  return String(value || "").trim().toLowerCase();
}

function cleanText(value) {
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
  if (!statusMessage) {
    return;
  }
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function hasOneTouchLink(candidate) {
  return Boolean(cleanText(candidate?.oneTouchLink));
}

function setAddButtonsBusy(disabled) {
  addToOneTouchBusy = disabled;
}

function setImportBusy(disabled) {
  importBusy = disabled;
  if (runImportBtn) {
    runImportBtn.disabled = disabled || pendingImportRows.length === 0;
  }
  if (importDropZone) {
    importDropZone.classList.toggle("is-disabled", disabled);
  }
}

function setImportSummary(message, isError = false) {
  if (!importSummary) {
    return;
  }
  importSummary.textContent = message;
  importSummary.classList.toggle("error", isError);
}

function setImportErrors(errors = []) {
  if (!importErrors) {
    return;
  }
  const validErrors = Array.isArray(errors) ? errors.filter(Boolean).slice(0, 8) : [];
  if (!validErrors.length) {
    importErrors.hidden = true;
    importErrors.textContent = "";
    return;
  }
  importErrors.hidden = false;
  importErrors.textContent = validErrors.join(" | ");
}

function stripTrailingUkPostcode(value) {
  const raw = cleanText(value);
  if (!raw) {
    return "";
  }
  const normalized = raw.replace(/\s+/g, " ").trim();
  const trimmed = normalized.replace(/[\s,;-]+([A-Z]{1,2}\d[A-Z\d]{0,2})$/i, "").trim();
  return trimmed || normalized;
}

function ensureIndeedPrefix(value) {
  const raw = cleanText(value);
  if (!raw) {
    return "Indeed";
  }
  const withoutIndeed = raw.replace(/^indeed(?:\s*[-:|]\s*|\s+)/i, "").trim();
  if (!withoutIndeed) {
    return "Indeed";
  }
  return `Indeed - ${withoutIndeed}`;
}

function normalizeImportStatus(statusValue, interestLevelValue) {
  const status = cleanText(statusValue);
  const interest = cleanText(interestLevelValue);
  const merged = `${status} ${interest}`.trim().toLowerCase();
  if (!merged) {
    return "Initial Call";
  }
  if (/\b(contacting|applied|application|new)\b/.test(merged)) {
    return "Initial Call";
  }
  if (/\b(interview|screening|screen)\b/.test(merged)) {
    return "1st Interview";
  }
  if (/\boffer\b/.test(merged)) {
    return "Offered";
  }
  if (/\b(hired|accepted)\b/.test(merged)) {
    return "Accepted";
  }
  if (/\brejected\b/.test(merged)) {
    return "Rejected";
  }
  if (/\blost\b/.test(merged)) {
    return "Lost";
  }
  return status || "Initial Call";
}

function getCsvValue(row, key) {
  const target = normalizeText(key);
  for (const [field, value] of Object.entries(row || {})) {
    if (normalizeText(field) === target) {
      return cleanText(value);
    }
  }
  return "";
}

function toImportPreviewRow(row) {
  return {
    candidateName: getCsvValue(row, "name"),
    email: getCsvValue(row, "email"),
    phone: getCsvValue(row, "phone"),
    livesIn: getCsvValue(row, "candidate location"),
    location: stripTrailingUkPostcode(getCsvValue(row, "job location")),
    status: normalizeImportStatus(getCsvValue(row, "status"), getCsvValue(row, "interest level")),
    source: ensureIndeedPrefix(getCsvValue(row, "source")),
  };
}

function renderImportPreview(rows) {
  if (!importPreviewWrap || !importPreviewBody || !importPreviewTitle) {
    return;
  }
  const list = Array.isArray(rows) ? rows : [];
  if (!list.length) {
    importPreviewWrap.hidden = true;
    importPreviewBody.innerHTML = "";
    importPreviewTitle.textContent = `Preview (top ${IMPORT_PREVIEW_LIMIT})`;
    return;
  }

  const previewRows = list.slice(0, IMPORT_PREVIEW_LIMIT).map(toImportPreviewRow);
  importPreviewWrap.hidden = false;
  importPreviewTitle.textContent = `Preview (top ${Math.min(IMPORT_PREVIEW_LIMIT, list.length)} of ${list.length})`;
  importPreviewBody.innerHTML = "";

  for (const row of previewRows) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(row.candidateName || "-")}</td>
      <td>${escapeHtml(row.email || "-")}</td>
      <td>${escapeHtml(row.phone || "-")}</td>
      <td>${escapeHtml(row.livesIn || "-")}</td>
      <td>${escapeHtml(row.location || "-")}</td>
      <td>${escapeHtml(row.status || "-")}</td>
      <td>${escapeHtml(row.source || "-")}</td>
    `;
    importPreviewBody.appendChild(tr);
  }
}

function formatBoolean(value) {
  return value === true ? "Yes" : "No";
}

function formatDate(value) {
  const raw = cleanText(value);
  if (!raw) {
    return "-";
  }
  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) {
    return raw;
  }
  return parsed.toLocaleDateString();
}

function setLinkField(node, url) {
  if (!node) {
    return;
  }
  const cleanUrl = cleanText(url);
  if (!cleanUrl) {
    node.textContent = "-";
    return;
  }
  const safeUrl = escapeHtml(cleanUrl);
  node.innerHTML = `<a href="${safeUrl}" target="_blank" rel="noopener noreferrer">Open link</a>`;
}

function setDetail(candidate) {
  if (!candidate) {
    detailFields.candidateName.textContent = "Select a candidate";
    detailFields.location.textContent = "-";
    detailFields.status.textContent = "-";
    detailFields.source.textContent = "-";
    detailFields.phoneNumber.textContent = "-";
    detailFields.interviewBooked.textContent = "-";
    detailFields.interviewWith.textContent = "-";
    detailFields.keepInMind.textContent = "-";
    detailFields.livesIn.textContent = "-";
    detailFields.firstInterviewDate.textContent = "-";
    detailFields.earmarkedFor.textContent = "-";
    detailFields.created.textContent = "-";
    detailFields.oneTouchLink.textContent = "-";
    detailFields.notes.textContent = "-";
    return;
  }

  detailFields.candidateName.textContent = cleanText(candidate.candidateName) || "-";
  detailFields.location.textContent = cleanText(candidate.location) || "-";
  detailFields.status.textContent = cleanText(candidate.status) || "-";
  detailFields.source.textContent = cleanText(candidate.source) || "-";
  detailFields.phoneNumber.textContent = cleanText(candidate.phoneNumber) || "-";
  detailFields.interviewBooked.textContent = formatBoolean(candidate.interviewBooked);
  detailFields.interviewWith.textContent = cleanText(candidate.interviewWith) || "-";
  detailFields.keepInMind.textContent = formatBoolean(candidate.keepInMind);
  detailFields.livesIn.textContent = cleanText(candidate.livesIn) || "-";
  detailFields.firstInterviewDate.textContent = formatDate(candidate.firstInterviewDate);
  detailFields.earmarkedFor.textContent = cleanText(candidate.earmarkedFor) || "-";
  detailFields.created.textContent = formatDate(candidate.created);
  setLinkField(detailFields.oneTouchLink, candidate.oneTouchLink);
  detailFields.notes.textContent = cleanText(candidate.notes) || "-";
}

function renderFilterOptions() {
  const locationOptions = Array.from(
    new Set(allCandidates.map((candidate) => cleanText(candidate.location)).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b));
  const statusOptions = Array.from(
    new Set(allCandidates.map((candidate) => cleanText(candidate.status)).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b));
  const sourceOptions = Array.from(
    new Set(allCandidates.map((candidate) => cleanText(candidate.source)).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b));

  const selectedLocation = cleanText(locationFilterSelect.value || "all");
  const selectedStatus = cleanText(statusFilterSelect.value || "all");
  const selectedSource = cleanText(sourceFilterSelect.value || "all");

  locationFilterSelect.innerHTML = '<option value="all">All locations</option>';
  statusFilterSelect.innerHTML = '<option value="all">All statuses</option>';
  sourceFilterSelect.innerHTML = '<option value="all">All sources</option>';

  for (const location of locationOptions) {
    const option = document.createElement("option");
    option.value = location;
    option.textContent = location;
    locationFilterSelect.appendChild(option);
  }
  for (const status of statusOptions) {
    const option = document.createElement("option");
    option.value = status;
    option.textContent = status;
    statusFilterSelect.appendChild(option);
  }
  for (const source of sourceOptions) {
    const option = document.createElement("option");
    option.value = source;
    option.textContent = source;
    sourceFilterSelect.appendChild(option);
  }

  locationFilterSelect.value = locationOptions.includes(selectedLocation) ? selectedLocation : "all";
  statusFilterSelect.value = statusOptions.includes(selectedStatus) ? selectedStatus : "all";
  sourceFilterSelect.value = sourceOptions.includes(selectedSource) ? selectedSource : "all";
}

function getFilteredCandidates() {
  const query = normalizeText(searchInput.value);
  const selectedLocation = cleanText(locationFilterSelect.value || "all");
  const selectedStatus = cleanText(statusFilterSelect.value || "all");
  const selectedSource = cleanText(sourceFilterSelect.value || "all");

  return allCandidates.filter((candidate) => {
    if (selectedLocation !== "all" && cleanText(candidate.location) !== selectedLocation) {
      return false;
    }
    if (selectedStatus !== "all" && cleanText(candidate.status) !== selectedStatus) {
      return false;
    }
    if (selectedSource !== "all" && cleanText(candidate.source) !== selectedSource) {
      return false;
    }
    if (!query) {
      return true;
    }
    return (
      normalizeText(candidate.candidateName).includes(query) ||
      normalizeText(candidate.location).includes(query) ||
      normalizeText(candidate.status).includes(query) ||
      normalizeText(candidate.source).includes(query) ||
      normalizeText(candidate.phoneNumber).includes(query) ||
      normalizeText(candidate.livesIn).includes(query) ||
      normalizeText(candidate.notes).includes(query)
    );
  });
}

function renderCandidates() {
  const filtered = getFilteredCandidates();
  recruitmentTableBody.innerHTML = "";

  if (!filtered.length) {
    emptyState.hidden = false;
    setDetail(null);
    return;
  }

  emptyState.hidden = true;
  const selected = filtered.find((candidate) => candidate.id === selectedCandidateId) || filtered[0];
  selectedCandidateId = selected.id;

  for (const candidate of filtered) {
    const tr = document.createElement("tr");
    tr.classList.toggle("selected", candidate.id === selectedCandidateId);
    tr.innerHTML = `
      <td>${escapeHtml(cleanText(candidate.candidateName) || "-")}</td>
      <td>${escapeHtml(cleanText(candidate.location) || "-")}</td>
      <td>${escapeHtml(cleanText(candidate.status) || "-")}</td>
      <td>${escapeHtml(cleanText(candidate.source) || "-")}</td>
      <td>${escapeHtml(cleanText(candidate.phoneNumber) || "-")}</td>
      <td>
        ${
          hasOneTouchLink(candidate)
            ? '<span class="muted">Added</span>'
            : `<button type="button" class="secondary recruitment-add-btn"${addToOneTouchBusy ? " disabled" : ""}>Add to OneTouch</button>`
        }
      </td>
    `;

    tr.addEventListener("click", () => {
      selectedCandidateId = candidate.id;
      setDetail(candidate);
      renderCandidates();
    });

    const addBtn = tr.querySelector(".recruitment-add-btn");
    addBtn?.addEventListener("click", async (event) => {
      event.stopPropagation();
      if (addToOneTouchBusy) {
        return;
      }
      await addCandidateToOneTouch(candidate.id);
    });

    recruitmentTableBody.appendChild(tr);
  }

  setDetail(selected);
}

function parseCsvLine(line) {
  const fields = [];
  let value = "";
  let inQuotes = false;

  for (let index = 0; index < line.length; index += 1) {
    const char = line[index];
    const next = line[index + 1];

    if (char === '"' && inQuotes && next === '"') {
      value += '"';
      index += 1;
      continue;
    }
    if (char === '"') {
      inQuotes = !inQuotes;
      continue;
    }
    if (char === "," && !inQuotes) {
      fields.push(value);
      value = "";
      continue;
    }
    value += char;
  }
  fields.push(value);
  return fields;
}

function parseCsvText(text) {
  const raw = String(text || "");
  const lines = raw.replace(/^\uFEFF/, "").split(/\r?\n/).filter((line) => line.trim() !== "");
  if (!lines.length) {
    return { rows: [], errors: ["CSV file is empty."] };
  }

  const headers = parseCsvLine(lines[0]).map((header) => cleanText(header));
  if (!headers.length) {
    return { rows: [], errors: ["CSV headers could not be read."] };
  }

  const rows = [];
  const errors = [];
  for (let lineIndex = 1; lineIndex < lines.length; lineIndex += 1) {
    const values = parseCsvLine(lines[lineIndex]);
    const row = {};
    for (let i = 0; i < headers.length; i += 1) {
      row[headers[i]] = cleanText(values[i] || "");
    }

    const isEmptyRow = Object.values(row).every((value) => !cleanText(value));
    if (isEmptyRow) {
      continue;
    }

    if (values.length !== headers.length) {
      errors.push(`Row ${lineIndex + 1} has ${values.length} value(s); expected ${headers.length}.`);
    }

    rows.push(row);
  }

  return { rows, errors };
}

async function previewImportRows(rows) {
  setImportBusy(true);
  setImportErrors([]);
  try {
    const preview = await directoryApi.previewRecruitmentImport({ rows });
    const summary = [
      `Rows: ${preview.totalRows}`,
      `Would insert: ${preview.wouldInsert}`,
      `Duplicates: ${preview.skippedDuplicates}`,
      `Rejected: ${preview.rejected}`,
    ].join(" | ");
    setImportSummary(summary);
    setImportErrors((preview.errors || []).map((error) => `Row ${error.row}: ${error.message}`));
    pendingImportRows = preview.wouldInsert > 0 ? rows : [];
  } catch (error) {
    pendingImportRows = [];
    setImportSummary(error?.message || "Could not preview CSV import.", true);
    setImportErrors([]);
  } finally {
    setImportBusy(false);
  }
}

async function handleCsvFile(file) {
  if (!file) {
    return;
  }
  if (importFileName) {
    importFileName.textContent = `Selected: ${file.name}`;
  }

  const text = await file.text();
  const parsed = parseCsvText(text);
  if (!parsed.rows.length) {
    pendingImportRows = [];
    renderImportPreview([]);
    setImportSummary("No importable rows found.", true);
    setImportErrors(parsed.errors);
    if (runImportBtn) {
      runImportBtn.disabled = true;
    }
    return;
  }

  if (parsed.errors.length) {
    setImportErrors(parsed.errors);
  } else {
    setImportErrors([]);
  }

  renderImportPreview(parsed.rows);
  await previewImportRows(parsed.rows);
}

async function runCsvImport() {
  if (importBusy || !pendingImportRows.length) {
    return;
  }
  setImportBusy(true);
  setImportErrors([]);
  try {
    const result = await directoryApi.runRecruitmentImport({ rows: pendingImportRows });
    const summary = [
      `Imported: ${result.inserted}`,
      `Duplicates skipped: ${result.skippedDuplicates}`,
      `Rejected: ${result.rejected}`,
    ].join(" | ");
    setImportSummary(summary);
    setImportErrors((result.errors || []).map((error) => `Row ${error.row}: ${error.message}`));
    pendingImportRows = [];
    renderImportPreview([]);
    if (importFileName) {
      importFileName.textContent = "No file selected.";
    }
    if (importFileInput) {
      importFileInput.value = "";
    }
    await loadRecruitmentCandidates();
    setStatus(`CSV import complete. ${summary}`);
  } catch (error) {
    console.error(error);
    setImportSummary(error?.message || "CSV import failed.", true);
  } finally {
    setImportBusy(false);
  }
}

function upsertCandidateInCache(item) {
  if (!item || !item.id) {
    return;
  }
  const index = allCandidates.findIndex((candidate) => candidate.id === item.id);
  if (index < 0) {
    allCandidates.push(item);
    return;
  }
  allCandidates[index] = item;
}

async function addCandidateToOneTouch(itemId) {
  const cleanItemId = cleanText(itemId);
  if (!cleanItemId) {
    return;
  }
  setAddButtonsBusy(true);
  try {
    setStatus("Adding candidate to OneTouch...");
    const result = await directoryApi.addRecruitmentCandidateToOneTouch({
      itemId: cleanItemId,
    });
    if (result?.item) {
      upsertCandidateInCache(result.item);
    }
    renderCandidates();
    setStatus(`Candidate added to OneTouch (ID: ${cleanText(result?.oneTouchId) || "-"})`);
  } catch (error) {
    console.error(error);
    setStatus(error?.message || "Could not add candidate to OneTouch.", true);
  } finally {
    setAddButtonsBusy(false);
  }
}

function redirectToUnauthorized(pageKey) {
  const page = encodeURIComponent(String(pageKey || "recruitment").trim().toLowerCase());
  window.location.href = `./unauthorized.html?page=${page}`;
}

async function loadRecruitmentCandidates() {
  const payload = await directoryApi.listRecruitment();
  allCandidates = Array.isArray(payload?.items) ? payload.items : [];
  if (sharePointListLink) {
    sharePointListLink.href = cleanText(payload?.listUrl) || "#";
  }
  renderFilterOptions();
  renderCandidates();
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
    if (!canAccessPage(role, "recruitment")) {
      redirectToUnauthorized("recruitment");
      return;
    }

    renderTopNavigation({ role });
    setStatus("Loading active candidates...");
    await loadRecruitmentCandidates();
    setStatus(`Loaded ${allCandidates.length} active candidate(s).`);
  } catch (error) {
    if (error?.status === 403) {
      redirectToUnauthorized("recruitment");
      return;
    }
    console.error(error);
    setStatus(error?.message || "Could not load recruitment candidates.", true);
    emptyState.hidden = false;
    setDetail(null);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

searchInput?.addEventListener("input", renderCandidates);
locationFilterSelect?.addEventListener("change", renderCandidates);
statusFilterSelect?.addEventListener("change", renderCandidates);
sourceFilterSelect?.addEventListener("change", renderCandidates);

importDropZone?.addEventListener("click", () => {
  if (importBusy) {
    return;
  }
  importFileInput?.click();
});

importDropZone?.addEventListener("keydown", (event) => {
  if (importBusy) {
    return;
  }
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    importFileInput?.click();
  }
});

importDropZone?.addEventListener("dragover", (event) => {
  event.preventDefault();
  if (!importBusy) {
    importDropZone.classList.add("is-dragover");
  }
});

importDropZone?.addEventListener("dragleave", () => {
  importDropZone.classList.remove("is-dragover");
});

importDropZone?.addEventListener("drop", async (event) => {
  event.preventDefault();
  importDropZone.classList.remove("is-dragover");
  if (importBusy) {
    return;
  }
  const file = event.dataTransfer?.files?.[0] || null;
  if (!file) {
    return;
  }
  await handleCsvFile(file);
});

importFileInput?.addEventListener("change", async () => {
  const file = importFileInput.files?.[0] || null;
  if (!file || importBusy) {
    return;
  }
  await handleCsvFile(file);
});

runImportBtn?.addEventListener("click", async () => {
  await runCsvImport();
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
