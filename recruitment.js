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
const oneTouchPickerModal = document.getElementById("oneTouchPickerModal");
const oneTouchPickerCandidate = document.getElementById("oneTouchPickerCandidate");
const oneTouchAreaSelect = document.getElementById("oneTouchAreaSelect");
const oneTouchRecruitmentSourceSelect = document.getElementById("oneTouchRecruitmentSourceSelect");
const oneTouchPickerError = document.getElementById("oneTouchPickerError");
const oneTouchPickerConfirmBtn = document.getElementById("oneTouchPickerConfirmBtn");
const oneTouchPickerCancelBtn = document.getElementById("oneTouchPickerCancelBtn");

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
let latestImportWouldInsert = 0;
let importEditingRowIndex = -1;
let importEditingDraft = null;
let oneTouchOptionsCache = null;
let oneTouchPickerCandidateId = "";

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
  if (oneTouchPickerConfirmBtn) {
    oneTouchPickerConfirmBtn.disabled = disabled;
  }
  if (oneTouchPickerCancelBtn) {
    oneTouchPickerCancelBtn.disabled = disabled;
  }
}

function setOneTouchPickerError(message = "") {
  if (!oneTouchPickerError) {
    return;
  }
  const text = cleanText(message);
  oneTouchPickerError.hidden = !text;
  oneTouchPickerError.textContent = text;
}

function closeOneTouchPicker() {
  oneTouchPickerCandidateId = "";
  if (oneTouchPickerModal) {
    oneTouchPickerModal.hidden = true;
  }
  setOneTouchPickerError("");
}

function normalizeToken(value) {
  return cleanText(value)
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function pickBestOption(options, hint) {
  const tokenHint = normalizeToken(hint);
  if (!tokenHint || !Array.isArray(options)) {
    return "";
  }
  const exact = options.find((option) => normalizeToken(option) === tokenHint);
  if (exact) {
    return exact;
  }
  const contains = options.find((option) => normalizeToken(option).includes(tokenHint));
  if (contains) {
    return contains;
  }
  const reverseContains = options.find((option) => tokenHint.includes(normalizeToken(option)));
  return reverseContains || "";
}

async function ensureOneTouchOptionsLoaded() {
  if (oneTouchOptionsCache) {
    return oneTouchOptionsCache;
  }
  const options = await directoryApi.getRecruitmentOneTouchOptions();
  oneTouchOptionsCache = {
    areas: Array.isArray(options?.areas) ? options.areas : [],
    recruitmentSources: Array.isArray(options?.recruitmentSources) ? options.recruitmentSources : [],
  };
  return oneTouchOptionsCache;
}

async function openOneTouchPicker(candidate) {
  const candidateId = cleanText(candidate?.id);
  if (!candidateId || !oneTouchPickerModal) {
    return;
  }
  oneTouchPickerCandidateId = candidateId;
  oneTouchPickerModal.hidden = false;
  setOneTouchPickerError("");
  if (oneTouchPickerCandidate) {
    oneTouchPickerCandidate.textContent = `Candidate: ${cleanText(candidate?.candidateName) || "-"}`;
  }

  try {
    oneTouchPickerConfirmBtn.disabled = true;
    oneTouchPickerCancelBtn.disabled = true;
    const options = await ensureOneTouchOptionsLoaded();

    if (oneTouchAreaSelect) {
      oneTouchAreaSelect.innerHTML = '<option value="">Select area</option>';
      for (const area of options.areas) {
        const option = document.createElement("option");
        option.value = area;
        option.textContent = area;
        oneTouchAreaSelect.appendChild(option);
      }
      oneTouchAreaSelect.value = pickBestOption(options.areas, candidate?.earmarkedFor) || "";
    }
    if (oneTouchRecruitmentSourceSelect) {
      oneTouchRecruitmentSourceSelect.innerHTML = '<option value="">Select recruitment source</option>';
      for (const source of options.recruitmentSources) {
        const option = document.createElement("option");
        option.value = source;
        option.textContent = source;
        oneTouchRecruitmentSourceSelect.appendChild(option);
      }
      oneTouchRecruitmentSourceSelect.value = pickBestOption(options.recruitmentSources, candidate?.source) || "";
    }
  } catch (error) {
    setOneTouchPickerError(error?.message || "Could not load OneTouch options.");
  } finally {
    oneTouchPickerConfirmBtn.disabled = false;
    oneTouchPickerCancelBtn.disabled = false;
  }
}

function updateRunImportButtonState() {
  if (!runImportBtn) {
    return;
  }
  runImportBtn.disabled =
    importBusy || pendingImportRows.length === 0 || latestImportWouldInsert <= 0 || importEditingRowIndex >= 0;
}

function setImportBusy(disabled) {
  importBusy = disabled;
  updateRunImportButtonState();
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

function toTitleCaseName(value) {
  const raw = cleanText(value).toLowerCase();
  if (!raw) {
    return "";
  }

  function capitalizeToken(token) {
    const clean = cleanText(token);
    if (!clean) {
      return "";
    }
    return clean
      .split("-")
      .map((part) => (part ? part.charAt(0).toUpperCase() + part.slice(1) : ""))
      .join("-");
  }

  return raw
    .split(/\s+/)
    .filter(Boolean)
    .map((word) => capitalizeToken(word))
    .join(" ");
}

function sanitizePhone(value) {
  const raw = cleanText(value);
  if (!raw) {
    return "";
  }
  return raw.replace(/^[\s'"`´‘’“”]+/, "").trim();
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

function setCsvValue(row, key, value) {
  const target = normalizeText(key);
  const cleanValue = cleanText(value);
  for (const existingKey of Object.keys(row || {})) {
    if (normalizeText(existingKey) === target) {
      row[existingKey] = cleanValue;
      return;
    }
  }
  row[key] = cleanValue;
}

function createImportEditDraft(row) {
  return {
    name: toTitleCaseName(getCsvValue(row, "name")),
    email: getCsvValue(row, "email"),
    phone: sanitizePhone(getCsvValue(row, "phone")),
    candidateLocation: getCsvValue(row, "candidate location"),
    jobLocation: getCsvValue(row, "job location"),
    status: getCsvValue(row, "status"),
    interestLevel: getCsvValue(row, "interest level"),
    source: getCsvValue(row, "source"),
  };
}

function toImportPreviewRow(row) {
  return {
    candidateName: toTitleCaseName(getCsvValue(row, "name")),
    email: getCsvValue(row, "email"),
    phone: sanitizePhone(getCsvValue(row, "phone")),
    candidateLocation: getCsvValue(row, "candidate location"),
    jobLocation: getCsvValue(row, "job location"),
    status: getCsvValue(row, "status"),
    interestLevel: getCsvValue(row, "interest level"),
    source: ensureIndeedPrefix(getCsvValue(row, "source")),
  };
}

function beginImportRowEdit(rowIndex, row) {
  importEditingRowIndex = rowIndex;
  importEditingDraft = createImportEditDraft(row);
  updateRunImportButtonState();
  renderImportPreview(pendingImportRows);
}

function cancelImportRowEdit() {
  importEditingRowIndex = -1;
  importEditingDraft = null;
  updateRunImportButtonState();
  renderImportPreview(pendingImportRows);
}

async function saveImportRowEdit(row) {
  if (!row || !importEditingDraft) {
    return;
  }

  setCsvValue(row, "name", toTitleCaseName(importEditingDraft.name));
  setCsvValue(row, "email", importEditingDraft.email);
  setCsvValue(row, "phone", sanitizePhone(importEditingDraft.phone));
  setCsvValue(row, "candidate location", importEditingDraft.candidateLocation);
  setCsvValue(row, "job location", importEditingDraft.jobLocation);
  setCsvValue(row, "status", importEditingDraft.status);
  setCsvValue(row, "interest level", importEditingDraft.interestLevel);
  setCsvValue(row, "source", importEditingDraft.source);

  importEditingRowIndex = -1;
  importEditingDraft = null;
  updateRunImportButtonState();
  renderImportPreview(pendingImportRows);
  await previewImportRows(pendingImportRows);
}

function renderImportPreview(rows) {
  if (!importPreviewWrap || !importPreviewBody || !importPreviewTitle) {
    return;
  }
  const list = Array.isArray(rows) ? rows : [];
  if (!list.length) {
    importPreviewWrap.hidden = true;
    importPreviewBody.innerHTML = "";
    importPreviewTitle.textContent = "Preview (all rows)";
    return;
  }

  const previewRows = list.map(toImportPreviewRow);
  importPreviewWrap.hidden = false;
  importPreviewTitle.textContent = `Preview (all ${list.length} rows)`;
  importPreviewBody.innerHTML = "";

  for (let rowIndex = 0; rowIndex < previewRows.length; rowIndex += 1) {
    const row = previewRows[rowIndex];
    const sourceRow = list[rowIndex];
    const tr = document.createElement("tr");
    const isEditing = importEditingRowIndex === rowIndex && importEditingDraft;
    tr.classList.toggle("import-preview-row", !isEditing);
    tr.classList.toggle("import-preview-row-editing", Boolean(isEditing));

    if (isEditing) {
      tr.innerHTML = `
        <td><input class="import-edit-input" data-field="name" type="text" value="${escapeHtml(importEditingDraft.name || "")}" /></td>
        <td><input class="import-edit-input" data-field="email" type="email" value="${escapeHtml(importEditingDraft.email || "")}" /></td>
        <td><input class="import-edit-input" data-field="phone" type="text" value="${escapeHtml(importEditingDraft.phone || "")}" /></td>
        <td><input class="import-edit-input" data-field="candidateLocation" type="text" value="${escapeHtml(importEditingDraft.candidateLocation || "")}" /></td>
        <td><input class="import-edit-input" data-field="jobLocation" type="text" value="${escapeHtml(importEditingDraft.jobLocation || "")}" /></td>
        <td><input class="import-edit-input" data-field="status" type="text" value="${escapeHtml(importEditingDraft.status || "")}" /></td>
        <td><input class="import-edit-input" data-field="interestLevel" type="text" value="${escapeHtml(importEditingDraft.interestLevel || "")}" /></td>
        <td><input class="import-edit-input" data-field="source" type="text" value="${escapeHtml(importEditingDraft.source || "")}" /></td>
      `;
      const inputs = tr.querySelectorAll(".import-edit-input");
      for (const input of inputs) {
        input.addEventListener("input", () => {
          const key = cleanText(input.getAttribute("data-field"));
          if (!key || !importEditingDraft) {
            return;
          }
          let nextValue = input.value;
          if (key === "name") {
            nextValue = toTitleCaseName(nextValue);
            if (input.value !== nextValue) {
              input.value = nextValue;
            }
          } else if (key === "phone") {
            nextValue = sanitizePhone(nextValue);
            if (input.value !== nextValue) {
              input.value = nextValue;
            }
          }
          importEditingDraft[key] = nextValue;
        });
      }
      const actionCell = document.createElement("td");
      actionCell.className = "import-preview-actions";
      actionCell.innerHTML = `
        <button type="button" class="secondary import-save-btn">Save</button>
        <button type="button" class="secondary import-cancel-btn">Cancel</button>
      `;
      actionCell.querySelector(".import-save-btn")?.addEventListener("click", async (event) => {
        event.stopPropagation();
        await saveImportRowEdit(sourceRow);
      });
      actionCell.querySelector(".import-cancel-btn")?.addEventListener("click", (event) => {
        event.stopPropagation();
        cancelImportRowEdit();
      });
      tr.appendChild(actionCell);
    } else {
      tr.innerHTML = `
        <td>${escapeHtml(row.candidateName || "-")}</td>
        <td>${escapeHtml(row.email || "-")}</td>
        <td>${escapeHtml(row.phone || "-")}</td>
        <td>${escapeHtml(row.candidateLocation || "-")}</td>
        <td>${escapeHtml(row.jobLocation || "-")}</td>
        <td>${escapeHtml(row.status || "-")}</td>
        <td>${escapeHtml(row.interestLevel || "-")}</td>
        <td>${escapeHtml(row.source || "-")}</td>
      `;
      const actionCell = document.createElement("td");
      actionCell.className = "import-preview-actions";
      actionCell.innerHTML = `<button type="button" class="secondary import-edit-btn">Edit</button>`;
      actionCell.querySelector(".import-edit-btn")?.addEventListener("click", (event) => {
        event.stopPropagation();
        beginImportRowEdit(rowIndex, sourceRow);
      });
      tr.appendChild(actionCell);
      tr.addEventListener("click", () => {
        beginImportRowEdit(rowIndex, sourceRow);
      });
    }
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
      await openOneTouchPicker(candidate);
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
    pendingImportRows = rows;
    latestImportWouldInsert = Number(preview.wouldInsert || 0);
    updateRunImportButtonState();
  } catch (error) {
    pendingImportRows = rows;
    latestImportWouldInsert = 0;
    updateRunImportButtonState();
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
    latestImportWouldInsert = 0;
    renderImportPreview([]);
    setImportSummary("No importable rows found.", true);
    setImportErrors(parsed.errors);
    updateRunImportButtonState();
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
    latestImportWouldInsert = 0;
    renderImportPreview([]);
    if (importFileName) {
      importFileName.textContent = "No file selected.";
    }
    if (importFileInput) {
      importFileInput.value = "";
    }
    updateRunImportButtonState();
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
  const selectedArea = cleanText(oneTouchAreaSelect?.value);
  const selectedRecruitmentSource = cleanText(oneTouchRecruitmentSourceSelect?.value);
  if (!selectedArea || !selectedRecruitmentSource) {
    setOneTouchPickerError("Select both area and recruitment source.");
    return;
  }

  setAddButtonsBusy(true);
  try {
    setStatus("Adding candidate to OneTouch...");
    const result = await directoryApi.addRecruitmentCandidateToOneTouch({
      itemId: cleanItemId,
      area: selectedArea,
      recruitmentSource: selectedRecruitmentSource,
    });
    if (result?.item) {
      upsertCandidateInCache(result.item);
    }
    renderCandidates();
    closeOneTouchPicker();
    setStatus(`Candidate added to OneTouch (ID: ${cleanText(result?.oneTouchId) || "-"})`);
  } catch (error) {
    console.error(error);
    setOneTouchPickerError(error?.message || "Could not add candidate to OneTouch.");
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

oneTouchPickerCancelBtn?.addEventListener("click", () => {
  if (addToOneTouchBusy) {
    return;
  }
  closeOneTouchPicker();
});

oneTouchPickerConfirmBtn?.addEventListener("click", async () => {
  if (addToOneTouchBusy) {
    return;
  }
  await addCandidateToOneTouch(oneTouchPickerCandidateId);
});

oneTouchPickerModal?.addEventListener("click", (event) => {
  if (event.target !== oneTouchPickerModal || addToOneTouchBusy) {
    return;
  }
  closeOneTouchPicker();
});

document.addEventListener("keydown", (event) => {
  if (event.key !== "Escape" || addToOneTouchBusy) {
    return;
  }
  if (!oneTouchPickerModal?.hidden) {
    closeOneTouchPicker();
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
