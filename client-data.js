import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260512";

const EXAMPLE_NOTE =
  "Paulette Crawley feels safe at Lindau. Carrie is brilliant. Discussed Martin Tyrell and selling her properties.";

const LLM_COPY_PREFIX = [
  "Instructions for the LLM:",
  "Keep every placeholder tag exactly as written, including the square brackets.",
  "Do not rename, renumber, remove, expand, or merge any text inside brackets such as [CLIENT_001] or [CLIENT_PREFERRED_NAME_001].",
  "In body text, use [CLIENT_PREFERRED_NAME_001] instead of [CLIENT_001] whenever the client is mentioned naturally in a sentence.",
  "Use [CLIENT_001] only for headings, labels, tables, or when the full placeholder is explicitly needed.",
  "If a preferred-name or surname variant is needed, convert [CLIENT_001] to [CLIENT_PREFERRED_NAME_001] or [CLIENT_SURNAME_001] with the same number.",
  "Do not invent new numbering. Keep the numeric suffix aligned with the original placeholder.",
  "You may rewrite the surrounding narrative, but preserve all bracketed tags verbatim.",
  "",
  "Text:",
].join("\n");

const PLACEHOLDER_CATEGORY_OPTIONS = [
  "CLIENT",
  "STAFF",
  "RELATIVE",
  "FRIEND",
  "PROFESSIONAL",
  "CARE_HOME",
  "LOCATION",
  "ORGANISATION",
  "DATE",
  "PHONE",
  "EMAIL",
  "ADDRESS",
  "IDENTIFIER",
];

const sourceText = document.getElementById("sourceText");
const reviewText = document.getElementById("reviewText");
const placeholderText = document.getElementById("placeholderText");
const restoredText = document.getElementById("restoredText");
const preferredNameInput = document.getElementById("preferredNameInput");
const manualIdentifierText = document.getElementById("manualIdentifierText");
const manualIdentifierCategory = document.getElementById("manualIdentifierCategory");
const addManualIdentifierBtn = document.getElementById("addManualIdentifierBtn");
const pseudonymiseBtn = document.getElementById("pseudonymiseBtn");
const clearBtn = document.getElementById("clearBtn");
const exampleBtn = document.getElementById("exampleBtn");
const copyLlmBtn = document.getElementById("copyLlmBtn");
const copyRestoredBtn = document.getElementById("copyRestoredBtn");
const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const riskBadge = document.getElementById("riskBadge");
const inputCount = document.getElementById("inputCount");
const outputCount = document.getElementById("outputCount");
const summaryPanel = document.getElementById("summaryPanel");
const summaryDetails = document.getElementById("summaryDetails");
const summaryMessage = document.getElementById("summaryMessage");
const countGrid = document.getElementById("countGrid");
const residualList = document.getElementById("residualList");
const mappingList = document.getElementById("mappingList");
const restoreStatus = document.getElementById("restoreStatus");

let account = null;
let result = null;
let manualMapping = {};
let busy = false;

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

function endpoint(pathname) {
  const base = String(FRONTEND_CONFIG.apiBaseUrl || "").replace(/\/+$/, "");
  return base ? `${base}${pathname}` : pathname;
}

function redirectToUnauthorized(pageKey) {
  const page = encodeURIComponent(String(pageKey || "clientdata").trim().toLowerCase());
  window.location.href = `./unauthorized.html?page=${page}`;
}

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function setRiskBadge(findings) {
  const riskLevel = String(findings?.riskLevel || "LOW").toLowerCase();
  riskBadge.className = `risk-badge risk-badge-${riskLevel}`;
  riskBadge.textContent = findings?.reason || "Ready";
}

function setBusy(nextBusy) {
  busy = Boolean(nextBusy);
  pseudonymiseBtn.disabled = busy || !sourceText.value.trim();
  pseudonymiseBtn.textContent = busy ? "Pseudonymising..." : "Pseudonymise";
}

async function apiPost(pathname, payload) {
  const token = await authController.acquireToken([FRONTEND_CONFIG.apiScope]);
  const response = await fetch(endpoint(pathname), {
    method: "POST",
    headers: {
      Accept: "application/json",
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });

  const text = await response.text();
  let data = null;
  try {
    data = text ? JSON.parse(text) : null;
  } catch {
    data = null;
  }

  if (!response.ok) {
    const detail = data?.error || text || `HTTP ${response.status}`;
    const error = new Error(String(detail));
    error.status = response.status;
    throw error;
  }

  return data || {};
}

async function copyTextToClipboard(text) {
  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(text);
    return;
  }

  const input = document.createElement("textarea");
  input.value = text;
  input.setAttribute("readonly", "");
  input.style.position = "fixed";
  input.style.opacity = "0";
  document.body.appendChild(input);
  input.select();
  const copied = document.execCommand("copy");
  document.body.removeChild(input);
  if (!copied) {
    throw new Error("Copy failed.");
  }
}

function characterLabel(length) {
  return `${Number(length || 0).toLocaleString()} character${length === 1 ? "" : "s"}`;
}

function updateCounts() {
  inputCount.textContent = characterLabel(sourceText.value.length);
  outputCount.textContent = characterLabel(reviewText.value.length);
  copyLlmBtn.disabled = !reviewText.value.trim();
  pseudonymiseBtn.disabled = busy || !sourceText.value.trim();
  updateManualControls();
  updateRestore();
}

function updateManualControls() {
  const value = manualIdentifierText.value.trim();
  addManualIdentifierBtn.disabled = !value || !reviewText.value.includes(value);
}

function clearAll() {
  sourceText.value = "";
  reviewText.value = "";
  placeholderText.value = "";
  restoredText.value = "";
  preferredNameInput.value = "";
  manualIdentifierText.value = "";
  result = null;
  manualMapping = {};
  if (summaryDetails) {
    delete summaryDetails.dataset.userOpened;
    summaryDetails.open = false;
  }
  setRiskBadge(null);
  setStatus("Paste a note to begin.");
  renderSummary();
  updateCounts();
}

function useExample() {
  sourceText.value = EXAMPLE_NOTE;
  setStatus("Example note loaded.");
  updateCounts();
}

async function pseudonymise() {
  const text = sourceText.value.trim();
  if (!text || busy) {
    return;
  }

  setBusy(true);
  setStatus("Pseudonymising note...");
  try {
    result = await apiPost("/api/pseudonymiser/pseudonymise", { text });
    manualMapping = {};
    if (summaryDetails) {
      delete summaryDetails.dataset.userOpened;
      summaryDetails.open = false;
    }
    reviewText.value = result.pseudonymised_text || "";
    if (!preferredNameInput.value.trim()) {
      preferredNameInput.value = defaultPreferredName(result.mapping || {});
    }
    setRiskBadge(result.findings);
    setStatus("Pseudonymised. Review highlighted residuals before copying.");
    renderSummary();
  } catch (error) {
    setStatus(error?.message || "Could not pseudonymise note.", true);
  } finally {
    setBusy(false);
    updateCounts();
  }
}

async function copyForLlm() {
  if (!reviewText.value.trim()) {
    return;
  }
  await copyTextToClipboard(`${LLM_COPY_PREFIX}\n${reviewText.value}`);
  setStatus("Copied pseudonymised text and placeholder instructions.");
}

async function copyRestored() {
  if (!restoredText.value.trim()) {
    return;
  }
  await copyTextToClipboard(restoredText.value);
  setStatus("Copied restored text.");
}

function addManualIdentifier() {
  const value = manualIdentifierText.value.trim();
  if (!value || !reviewText.value.includes(value)) {
    updateManualControls();
    return;
  }

  const mapping = buildEffectiveMapping();
  const placeholder = nextManualPlaceholder(manualIdentifierCategory.value, mapping);
  manualMapping = { ...manualMapping, [placeholder]: value };
  reviewText.value = reviewText.value.split(value).join(placeholder);
  manualIdentifierText.value = "";
  setStatus(`Added ${placeholder}.`);
  renderSummary();
  updateCounts();
}

function applyResidual(span) {
  const value = String(result?.pseudonymised_text || "").slice(span.start, span.end);
  if (!value || !reviewText.value.includes(value)) {
    return;
  }

  const mapping = buildEffectiveMapping();
  const placeholder = nextManualPlaceholder(span.category || "IDENTIFIER", mapping);
  manualMapping = { ...manualMapping, [placeholder]: value };
  reviewText.value = reviewText.value.split(value).join(placeholder);
  setStatus(`Replaced residual with ${placeholder}.`);
  renderSummary();
  updateCounts();
}

function buildEffectiveMapping() {
  const merged = { ...(result?.mapping || {}), ...manualMapping };
  const preferredClientName = normalisePreferredName(preferredNameInput.value);

  for (const [placeholder, original] of Object.entries({ ...merged })) {
    const match = placeholder.match(/^\[([A-Z_]+)_(\d{3})\]$/);
    if (!match) {
      continue;
    }

    const [, category, index] = match;
    const { firstName, surname } = splitPersonName(original);
    const preferredName = category === "CLIENT" ? preferredClientName || firstName : firstName;

    if (preferredName) {
      merged[`[${category}_PREFERRED_NAME_${index}]`] = preferredName;
      merged[`[${category}_FIRST_NAME_${index}]`] = preferredName;
    }
    if (surname) {
      merged[`[${category}_SURNAME_${index}]`] = surname;
    }
  }

  return merged;
}

function defaultPreferredName(mapping) {
  const clientName = mapping["[CLIENT_001]"];
  if (!clientName) {
    return "";
  }
  return splitPersonName(clientName).firstName || "";
}

function splitPersonName(value) {
  const cleaned = String(value || "")
    .trim()
    .replace(/\b(?:Mr|Mrs|Ms|Miss|Dr)\.?\s+/gi, "")
    .replace(/\s+/g, " ");
  const parts = cleaned.split(" ").filter(Boolean);
  if (parts.length === 0) {
    return { firstName: null, surname: null };
  }
  if (parts.length === 1) {
    return { firstName: parts[0], surname: null };
  }
  return { firstName: parts[0], surname: parts[parts.length - 1] };
}

function normalisePreferredName(value) {
  return String(value || "").trim().replace(/\s+/g, " ") || null;
}

function placeholderSafeCategory(category) {
  const clean = String(category || "")
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
  return clean || "IDENTIFIER";
}

function nextManualPlaceholder(category, mapping) {
  const cleanCategory = placeholderSafeCategory(category);
  const highest = Object.keys(mapping || {}).reduce((max, placeholder) => {
    const match = placeholder.match(/^\[([A-Z_]+)_(\d{3})\]$/);
    if (!match || match[1] !== cleanCategory) {
      return max;
    }
    return Math.max(max, Number(match[2]));
  }, 0);
  return `[${cleanCategory}_${String(highest + 1).padStart(3, "0")}]`;
}

function deanonymiseText(mapping, text) {
  let restoredCount = 0;
  let unresolvedCount = 0;
  const output = String(text || "").replace(/(?:\[[A-Z0-9_]+\]|<[A-Z0-9_]+>)/g, (placeholder) => {
    const value = resolvePlaceholderValue(mapping, placeholder);
    if (!value) {
      unresolvedCount += 1;
      return placeholder;
    }
    restoredCount += 1;
    return value;
  });
  return { text: output, restoredCount, unresolvedCount };
}

function resolvePlaceholderValue(mapping, placeholder) {
  if (mapping[placeholder]) {
    return mapping[placeholder];
  }

  const match = placeholder.match(/^[<\[]([A-Z_]+)_(PREFERRED_NAME|FIRST_NAME|SURNAME)_(\d{3})[>\]]$/);
  if (!match) {
    return null;
  }

  const [, category, part, index] = match;
  const baseValue = mapping[`[${category}_${index}]`] || mapping[`<${category}_${index}>`];
  if (!baseValue) {
    return null;
  }

  const { firstName, surname } = splitPersonName(baseValue);
  if (part === "SURNAME") {
    return surname;
  }

  return mapping[`[${category}_PREFERRED_NAME_${index}]`] || mapping[`<${category}_PREFERRED_NAME_${index}>`] || firstName;
}

function updateRestore() {
  const mapping = buildEffectiveMapping();
  const restored = deanonymiseText(mapping, placeholderText.value);
  restoredText.value = restored.text;
  copyRestoredBtn.disabled = !restored.text.trim();
  if (!placeholderText.value.trim()) {
    restoreStatus.textContent = "Paste placeholder-bearing output to restore with the local map.";
    return;
  }
  restoreStatus.textContent = `${restored.restoredCount.toLocaleString()} restored${
    restored.unresolvedCount > 0 ? `, ${restored.unresolvedCount.toLocaleString()} unresolved` : ""
  }.`;
}

function renderSummary() {
  summaryPanel.hidden = !result;
  if (result && summaryDetails && !summaryDetails.dataset.userOpened) {
    summaryDetails.open = false;
  }
  countGrid.innerHTML = "";
  residualList.innerHTML = "";
  mappingList.innerHTML = "";
  if (!result) {
    return;
  }

  const counts = Array.isArray(result.counts) && result.counts.length ? result.counts : [{ entity_type: "No findings", count: 0 }];
  for (const item of counts) {
    const card = document.createElement("div");
    card.className = "client-data-count";
    card.innerHTML = `<span>${escapeHtml(String(item.entity_type || "").replace(/_/g, " "))}</span><strong>${Number(item.count || 0).toLocaleString()}</strong>`;
    countGrid.appendChild(card);
  }

  const spans = (result.findings?.residualSpans || []).filter((span) => {
    const value = String(result.pseudonymised_text || "").slice(span.start, span.end);
    return value && reviewText.value.includes(value);
  });
  summaryMessage.textContent =
    spans.length > 0 ? "Review any residual identifiers before copying." : "No unresolved residual identifiers are currently visible.";

  if (!spans.length) {
    residualList.innerHTML = '<p class="muted">No visible residual identifiers.</p>';
  } else {
    for (const span of spans) {
      const value = String(result.pseudonymised_text || "").slice(span.start, span.end);
      const row = document.createElement("div");
      row.className = "client-data-list-row";
      row.innerHTML = `
        <div>
          <strong>${escapeHtml(value)}</strong>
          <span>${escapeHtml(span.reason || "Residual identifier")} · ${escapeHtml(span.severity || "review")}</span>
        </div>
      `;
      const button = document.createElement("button");
      button.className = "secondary";
      button.type = "button";
      button.textContent = "Replace";
      button.addEventListener("click", () => applyResidual(span));
      row.appendChild(button);
      residualList.appendChild(row);
    }
  }

  const mapping = buildEffectiveMapping();
  const baseEntries = Object.entries(mapping).filter(([placeholder]) => /^\[[A-Z_]+_\d{3}\]$/.test(placeholder));
  if (!baseEntries.length) {
    mappingList.innerHTML = '<p class="muted">No placeholders yet.</p>';
  } else {
    for (const [placeholder, original] of baseEntries) {
      const row = document.createElement("div");
      row.className = "client-data-list-row";
      row.innerHTML = `
        <div>
          <strong>${escapeHtml(placeholder)}</strong>
          <span>${escapeHtml(original)}</span>
        </div>
      `;
      mappingList.appendChild(row);
    }
  }

  updateRestore();
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
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
    if (!canAccessPage(role, "clientdata")) {
      redirectToUnauthorized("clientdata");
      return;
    }
    renderTopNavigation({ role });
    setStatus("Paste a note to begin.");
  } catch (error) {
    if (error?.status === 403) {
      redirectToUnauthorized("clientdata");
      return;
    }
    setStatus(error?.message || "Could not load client data page.", true);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

for (const category of PLACEHOLDER_CATEGORY_OPTIONS) {
  const option = document.createElement("option");
  option.value = category;
  option.textContent = category.replace(/_/g, " ");
  manualIdentifierCategory.appendChild(option);
}
manualIdentifierCategory.value = "PROFESSIONAL";

sourceText.addEventListener("input", updateCounts);
reviewText.addEventListener("input", () => {
  renderSummary();
  updateCounts();
});
placeholderText.addEventListener("input", updateRestore);
preferredNameInput.addEventListener("input", () => {
  renderSummary();
  updateRestore();
});
manualIdentifierText.addEventListener("input", updateManualControls);
pseudonymiseBtn.addEventListener("click", () => {
  void pseudonymise();
});
clearBtn.addEventListener("click", clearAll);
exampleBtn.addEventListener("click", useExample);
copyLlmBtn.addEventListener("click", () => {
  void copyForLlm();
});
copyRestoredBtn.addEventListener("click", () => {
  void copyRestored();
});
addManualIdentifierBtn.addEventListener("click", addManualIdentifier);
summaryDetails?.addEventListener("toggle", () => {
  if (summaryDetails.open) {
    summaryDetails.dataset.userOpened = "true";
  }
});
signOutBtn?.addEventListener("click", () => {
  authController.signOut({ redirectUri: `${window.location.origin}/index.html` }).catch(() => {});
});

updateCounts();
void init();
