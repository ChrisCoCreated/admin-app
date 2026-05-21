import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260512";

const EXAMPLE_NOTE =
  "Paul Jones feels safe at Mossbank. Claire is brilliant. Discussed Bob Smith and selling her properties.";

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

const REPORT_MODES = {
  wellbeing_assurance_visit: {
    label: "Wellbeing Assurance Visit Summary",
    reportType: "wellbeing_assurance_visit",
    defaultProvider: "azure_openai",
    defaultModel: "primary",
    modelLocked: true,
  },
  simple_summary: {
    label: "Simple summary report",
    reportType: "simple_summary",
    defaultProvider: "azure_openai",
    defaultModel: "fast",
    modelLocked: false,
  },
};

const REPORT_MODEL_OPTIONS = {
  azure_openai: [
    { value: "primary", label: "OpenAI primary" },
    { value: "fast", label: "OpenAI fast" },
  ],
  deepseek: [
    { value: "deepseek-v4-flash", label: "DeepSeek V4 Flash" },
    { value: "deepseek-v4-pro", label: "DeepSeek V4 Pro" },
  ],
};

const sourceText = document.getElementById("sourceText");
const reviewOutput = document.getElementById("reviewOutput");
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
const reportStatus = document.getElementById("reportStatus");
const reportProviderSelect = document.getElementById("reportProviderSelect");
const reportModelSelect = document.getElementById("reportModelSelect");
const reportThinkingField = document.getElementById("reportThinkingField");
const reportThinkingSelect = document.getElementById("reportThinkingSelect");
const reportModeSelect = document.getElementById("reportModeSelect");
const reportModelLockBtn = document.getElementById("reportModelLockBtn");
const generateReportBtn = document.getElementById("generateReportBtn");
const reportRevisionPanel = document.getElementById("reportRevisionPanel");
const reportRevisionText = document.getElementById("reportRevisionText");
const reviseReportBtn = document.getElementById("reviseReportBtn");
const exportStatus = document.getElementById("exportStatus");
const exportPdfBtn = document.getElementById("exportPdfBtn");
const exportWordBtn = document.getElementById("exportWordBtn");

let account = null;
let result = null;
let reviewTextValue = "";
let manualMapping = {};
let busy = false;
let reportBusy = false;
let structuredReport = null;
let reportSourceNotes = "";
let reportModelUnlocked = false;

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
  updateReportControls();
}

function setReportStatus(message, isError = false) {
  if (!reportStatus) {
    return;
  }
  reportStatus.textContent = message;
  reportStatus.classList.toggle("error", isError);
}

function setReportBusy(nextBusy) {
  reportBusy = Boolean(nextBusy);
  if (generateReportBtn) {
    generateReportBtn.textContent = reportBusy ? "Generating..." : "Generate report";
  }
  updateReportControls();
}

function getSelectedReportProvider() {
  return reportProviderSelect?.value || "azure_openai";
}

function getSelectedReportMode() {
  return REPORT_MODES[reportModeSelect?.value] || REPORT_MODES.wellbeing_assurance_visit;
}

function getSelectedReportType() {
  const mode = getSelectedReportMode();
  return mode.reportType || "wellbeing_assurance_visit";
}

function getGeneratedReportType() {
  return structuredReport?.report_type || getSelectedReportType();
}

function syncReportThinkingControl() {
  const isAzure = getSelectedReportProvider() === "azure_openai";
  if (reportThinkingField) {
    reportThinkingField.hidden = isAzure;
  }
  if (isAzure && reportThinkingSelect) {
    reportThinkingSelect.value = "disabled";
  }
}

function isReportModelLocked() {
  return Boolean(getSelectedReportMode().modelLocked && !reportModelUnlocked);
}

function updateReportModelLockControl() {
  const mode = getSelectedReportMode();
  const lockApplies = Boolean(mode.modelLocked);
  const locked = isReportModelLocked();
  if (reportProviderSelect) {
    reportProviderSelect.disabled = reportBusy || locked;
  }
  if (reportModelSelect) {
    reportModelSelect.disabled = reportBusy || locked;
  }
  if (reportModelLockBtn) {
    reportModelLockBtn.hidden = !lockApplies;
    reportModelLockBtn.parentElement?.classList.toggle("is-lock-hidden", !lockApplies);
    reportModelLockBtn.disabled = reportBusy || !lockApplies;
    reportModelLockBtn.classList.toggle("is-unlocked", lockApplies && !locked);
    reportModelLockBtn.setAttribute(
      "aria-label",
      locked ? "Unlock report model controls" : "Lock report model controls"
    );
    reportModelLockBtn.title = locked ? "Unlock report model controls" : "Lock report model controls";
  }
}

function getSelectedThinkingOptions() {
  if (getSelectedReportProvider() === "azure_openai") {
    return {
      thinking: "disabled",
      reasoningEffort: null,
    };
  }

  const value = reportThinkingSelect?.value || "disabled";
  if (value === "disabled") {
    return {
      thinking: "disabled",
      reasoningEffort: null,
    };
  }
  return {
    thinking: "enabled",
    reasoningEffort: value,
  };
}

function updateRevisionControls() {
  const canRevise = Boolean(structuredReport && structuredReport.status === "ready_for_render");
  if (reportRevisionPanel) {
    reportRevisionPanel.hidden = !canRevise;
  }
  if (reviseReportBtn) {
    reviseReportBtn.disabled = reportBusy || !canRevise || !reportRevisionText?.value.trim();
  }
}

function populateReportModels(selectedModel = "") {
  if (!reportModelSelect) {
    return;
  }

  const provider = getSelectedReportProvider();
  const options = REPORT_MODEL_OPTIONS[provider] || REPORT_MODEL_OPTIONS.azure_openai;
  const previousValue = reportModelSelect.value;
  const preferredValue = selectedModel || previousValue;
  reportModelSelect.innerHTML = "";

  for (const item of options) {
    const option = document.createElement("option");
    option.value = item.value;
    option.textContent = item.label;
    reportModelSelect.appendChild(option);
  }

  if (options.some((item) => item.value === preferredValue)) {
    reportModelSelect.value = preferredValue;
  }
}

function applyReportModeDefaults() {
  const mode = getSelectedReportMode();
  reportModelUnlocked = false;
  if (reportProviderSelect && mode.defaultProvider) {
    reportProviderSelect.value = mode.defaultProvider;
  }
  populateReportModels(mode.defaultModel);
  syncReportThinkingControl();
  updateReportControls();
}

function setExportStatus(message, isError = false) {
  if (!exportStatus) {
    return;
  }
  exportStatus.textContent = message;
  exportStatus.classList.toggle("error", isError);
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
  outputCount.textContent = characterLabel(reviewTextValue.length);
  copyLlmBtn.disabled = !reviewTextValue.trim();
  pseudonymiseBtn.disabled = busy || !sourceText.value.trim();
  updateManualControls();
  updateReportControls();
  updateExportControls();
  updateRestore();
}

function updateManualControls() {
  const value = manualIdentifierText.value.trim();
  addManualIdentifierBtn.disabled = !value || !reviewTextValue.includes(value);
}

function updateReportControls() {
  if (!generateReportBtn) {
    return;
  }
  generateReportBtn.disabled = reportBusy || busy;
  updateReportModelLockControl();
  updateRevisionControls();
}

function updateExportControls() {
  const hasRestoredText = Boolean(restoredText.value.trim());
  const canExport = hasRestoredText && (!structuredReport || structuredReport.status === "ready_for_render");
  if (exportPdfBtn) {
    exportPdfBtn.disabled = !canExport;
  }
  if (exportWordBtn) {
    exportWordBtn.disabled = !canExport || !structuredReport;
  }
}

function setReviewText(nextText) {
  reviewTextValue = String(nextText || "");
  renderReviewOutput();
  renderSummary();
  updateCounts();
}

function replaceReviewRange(start, end, value) {
  setReviewText(`${reviewTextValue.slice(0, start)}${value}${reviewTextValue.slice(end)}`);
}

function clearAll() {
  sourceText.value = "";
  reviewTextValue = "";
  placeholderText.value = "";
  restoredText.value = "";
  preferredNameInput.value = "";
  manualIdentifierText.value = "";
  result = null;
  manualMapping = {};
  structuredReport = null;
  reportSourceNotes = "";
  if (reportRevisionText) {
    reportRevisionText.value = "";
  }
  if (summaryDetails) {
    delete summaryDetails.dataset.userOpened;
    summaryDetails.open = false;
  }
  setRiskBadge(null);
  setStatus("Paste a note to begin.");
  setReportStatus("Generate a placeholder-preserving report from the pseudonymised text.");
  setExportStatus("Export the restored report text.");
  renderReviewOutput();
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
    setReviewText(result.pseudonymised_text || "");
    structuredReport = null;
    reportSourceNotes = "";
    if (reportRevisionText) {
      reportRevisionText.value = "";
    }
    if (!preferredNameInput.value.trim()) {
      preferredNameInput.value = defaultPreferredName(result.mapping || {});
    }
    renderReviewOutput();
    renderSummary();
    updateRestore();
    setRiskBadge(result.findings);
    setStatus("Pseudonymised. Review highlighted residuals before copying.");
    setReportStatus("Ready to generate a report from the pseudonymised text.");
  } catch (error) {
    setStatus(error?.message || "Could not pseudonymise note.", true);
  } finally {
    setBusy(false);
    updateCounts();
  }
}

async function copyForLlm() {
  if (!reviewTextValue.trim()) {
    return;
  }
  await copyTextToClipboard(`${LLM_COPY_PREFIX}\n${reviewTextValue}`);
  setStatus("Copied pseudonymised text and placeholder instructions.");
}

async function copyRestored() {
  if (!restoredText.value.trim()) {
    return;
  }
  await copyTextToClipboard(restoredText.value);
  setStatus("Copied restored text.");
}

function stringifyReportValue(value) {
  if (Array.isArray(value)) {
    return value.map(stringifyReportValue).filter(Boolean).join("; ");
  }
  if (value && typeof value === "object") {
    return Object.values(value).map(stringifyReportValue).filter(Boolean).join(" ");
  }
  return String(value || "").trim();
}

function formatReportList(items) {
  return (Array.isArray(items) ? items : [])
    .map((item) => stringifyReportValue(item))
    .filter(Boolean)
    .map((item) => `- ${item}`)
    .join("\n");
}

function buildClientDetailsText(clientDetails = {}) {
  const rows = [
    ["Client", clientDetails.client_name],
    ["Report date", clientDetails.report_date],
    ["Conducted by", clientDetails.conducted_by],
    ["Assessment venue", clientDetails.assessment_venue],
    ["Date of birth", clientDetails.date_of_birth],
    ["Prepared for", clientDetails.prepared_for],
  ];
  return rows
    .map(([label, value]) => {
      const text = stringifyReportValue(value);
      return text ? `${label}: ${text}` : "";
    })
    .filter(Boolean)
    .join("\n");
}

function buildStructuredReportText(report) {
  if (!report || typeof report !== "object") {
    return "";
  }
  const sections = report.report_sections || {};
  const omitted = new Set((report.omitted_sections || []).map((item) => String(item || "").toLowerCase()));
  const clientDetails = buildClientDetailsText(report.client_details);
  const parts = [
    report.report_title || buildExportTitle(),
    clientDetails,
    sections.executive_summary ? `Executive Summary\n${sections.executive_summary}` : "",
    sections.current_situation ? `1. Current Situation\n${sections.current_situation}` : "",
    sections.physical_wellbeing ? `2. Physical Wellbeing\n${sections.physical_wellbeing}` : "",
    sections.emotional_wellbeing ? `3. Emotional Wellbeing\n${sections.emotional_wellbeing}` : "",
    sections.environmental_wellbeing ? `4. Environmental Wellbeing\n${sections.environmental_wellbeing}` : "",
  ];
  if (sections.wellbeing_highlights && !omitted.has("wellbeing highlights")) {
    parts.push(`5. Wellbeing Highlights\n${sections.wellbeing_highlights}`);
  }
  const goals = formatReportList(report.suggested_smart_goals);
  if (goals) {
    parts.push(`Suggested SMART Goals\n${goals}\nThese goals are suggested and can be reviewed or amended.`);
  }
  const recommendations = formatReportList(sections.recommendations);
  if (recommendations) {
    parts.push(`Recommendations to Improve Quality of Life\n${recommendations}`);
  }
  const nextSteps = formatReportList(sections.next_steps);
  if (nextSteps) {
    parts.push(`Next Steps\n${nextSteps}`);
  }
  const warnings = formatReportList(report.warnings);
  if (warnings) {
    parts.push(`Warnings\n${warnings}`);
  }
  const clarifications = formatReportList(report.clarification_notes);
  if (clarifications) {
    parts.push(`Clarification Notes\n${clarifications}`);
  }
  return parts.filter(Boolean).join("\n\n");
}

function setStructuredReport(report) {
  structuredReport = report && typeof report === "object" ? report : null;
  if (!structuredReport) {
    placeholderText.value = "";
    updateRestore();
    updateRevisionControls();
    updateExportControls();
    return;
  }

  if (structuredReport.status === "needs_notes") {
    placeholderText.value = "";
    updateRestore();
    setReportStatus("Please provide notes.", true);
  } else if (structuredReport.status === "needs_clarification") {
    placeholderText.value = buildStructuredReportText(structuredReport);
    updateRestore();
    setReportStatus("Clarification needed before this report can be rendered.", true);
  } else {
    placeholderText.value = buildStructuredReportText(structuredReport);
    updateRestore();
    setReportStatus("Report generated. Preview restored text below and export as Word when ready.");
  }
  updateRevisionControls();
  updateExportControls();
}

async function generateReport(revisionRequest = "") {
  if (reportBusy) {
    return;
  }

  setReportBusy(true);
  const thinkingOptions = getSelectedThinkingOptions();
  const isRevision = Boolean(String(revisionRequest || "").trim() && structuredReport);
  setReportStatus(isRevision ? "Regenerating report..." : "Generating report...");
  try {
    const response = await directoryApi.generateStructuredReport({
      reportType: getSelectedReportType(),
      notes: isRevision ? reportSourceNotes || reviewTextValue : reviewTextValue,
      provider: getSelectedReportProvider(),
      model: reportModelSelect?.value,
      thinking: thinkingOptions.thinking,
      reasoningEffort: thinkingOptions.reasoningEffort,
      previousReport: isRevision ? structuredReport : null,
      revisionRequest: isRevision ? revisionRequest : "",
    });
    const report = response?.report;
    if (!report) {
      throw new Error("AI returned an empty report payload.");
    }
    reportSourceNotes = reviewTextValue;
    setStructuredReport(report);
    setStatus(report.status === "ready_for_render" ? "Generated structured report." : "Report needs more input.");
    if (isRevision && reportRevisionText) {
      reportRevisionText.value = "";
    }
  } catch (error) {
    const message = error?.message || "Could not generate report.";
    if (/Azure OpenAI deployment not found/i.test(message)) {
      setReportStatus(
        "Azure OpenAI deployment not found. Check the selected OpenAI model/deployment configuration.",
        true
      );
    } else {
      setReportStatus(message, true);
    }
  } finally {
    setReportBusy(false);
  }
}

async function reviseReport() {
  const request = reportRevisionText?.value.trim();
  if (!request || !structuredReport) {
    updateRevisionControls();
    return;
  }
  await generateReport(request);
}

function addManualIdentifier() {
  const value = manualIdentifierText.value.trim();
  if (!value || !reviewTextValue.includes(value)) {
    updateManualControls();
    return;
  }

  const mapping = buildEffectiveMapping();
  const placeholder = nextManualPlaceholder(manualIdentifierCategory.value, mapping);
  manualMapping = { ...manualMapping, [placeholder]: value };
  setReviewText(reviewTextValue.split(value).join(placeholder));
  manualIdentifierText.value = "";
  setStatus(`Added ${placeholder}.`);
}

function applyResidual(span) {
  const value = String(result?.pseudonymised_text || "").slice(span.start, span.end);
  if (!value || !reviewTextValue.includes(value)) {
    return;
  }

  const mapping = buildEffectiveMapping();
  const placeholder = nextManualPlaceholder(span.category || "IDENTIFIER", mapping);
  manualMapping = { ...manualMapping, [placeholder]: value };
  setReviewText(reviewTextValue.split(value).join(placeholder));
  setStatus(`Replaced residual with ${placeholder}.`);
}

function getReviewMarks() {
  const mapping = buildEffectiveMapping();
  const residualSpans = normaliseResidualSpans(result?.findings?.residualSpans || [], String(result?.pseudonymised_text || "").length);
  return buildReviewMarks(reviewTextValue, result?.pseudonymised_text || "", mapping, residualSpans);
}

function renderReviewOutput() {
  reviewOutput.innerHTML = "";
  if (!reviewTextValue) {
    reviewOutput.classList.add("client-data-review-output-placeholder");
    reviewOutput.textContent = "Pseudonymised output will appear here.";
    return;
  }

  reviewOutput.classList.remove("client-data-review-output-placeholder");
  const marks = getReviewMarks();
  const identifierOptions = buildIdentifierOptions(buildEffectiveMapping());
  let cursor = 0;

  for (const mark of marks) {
    if (mark.start > cursor) {
      reviewOutput.appendChild(document.createTextNode(reviewTextValue.slice(cursor, mark.start)));
    }
    reviewOutput.appendChild(renderReviewMark(mark, identifierOptions));
    cursor = mark.end;
  }

  if (cursor < reviewTextValue.length) {
    reviewOutput.appendChild(document.createTextNode(reviewTextValue.slice(cursor)));
  }
}

function getSelectedReviewText() {
  const selection = window.getSelection?.();
  if (!selection || selection.rangeCount === 0 || selection.isCollapsed) {
    return "";
  }

  const range = selection.getRangeAt(0);
  const container = range.commonAncestorContainer;
  const selectionRoot = container.nodeType === Node.ELEMENT_NODE ? container : container.parentElement;
  if (!selectionRoot || !reviewOutput.contains(selectionRoot)) {
    return "";
  }

  return selection.toString().trim();
}

function onApplySelectedReviewText() {
  const value = getSelectedReviewText();
  if (!value || !reviewTextValue.includes(value)) {
    return false;
  }
  if (/^\[[A-Z_]+_\d{3}\]$/.test(value)) {
    setStatus("Selected text is already a placeholder.");
    window.getSelection?.().removeAllRanges();
    return true;
  }

  const mapping = buildEffectiveMapping();
  const placeholder = nextManualPlaceholder(manualIdentifierCategory.value || "IDENTIFIER", mapping);
  manualMapping = { ...manualMapping, [placeholder]: value };
  setReviewText(reviewTextValue.split(value).join(placeholder));
  window.getSelection?.().removeAllRanges();
  setStatus(`Replaced all selected text matches with ${placeholder}.`);
  return true;
}

function renderReviewMark(mark, identifierOptions) {
  const content = reviewTextValue.slice(mark.start, mark.end);
  if (mark.kind === "replaced") {
    const wrap = document.createElement("span");
    wrap.className = "client-data-replacement";

    const button = document.createElement("button");
    button.className = "review-highlight replaced";
    button.type = "button";
    button.title = `Revert: ${mark.original}`;
    const placeholderTextSpan = document.createElement("span");
    placeholderTextSpan.className = "client-data-replacement-placeholder";
    placeholderTextSpan.textContent = content;
    button.appendChild(placeholderTextSpan);
    const originalTextSpan = document.createElement("span");
    originalTextSpan.className = "client-data-replacement-original";
    originalTextSpan.textContent = mark.original;
    button.appendChild(originalTextSpan);
    button.addEventListener("click", () => {
      replaceReviewRange(mark.start, mark.end, mark.original);
      setStatus(`Reverted ${mark.placeholder} for review.`);
    });
    wrap.appendChild(button);

    const select = document.createElement("select");
    select.className = "client-data-inline-select";
    select.value = mark.placeholder;
    select.setAttribute("aria-label", `Change identifier for ${mark.placeholder}`);
    for (const optionValue of identifierOptions) {
      const option = document.createElement("option");
      option.value = optionValue;
      option.textContent = identifierOptionLabel(optionValue);
      select.appendChild(option);
    }
    select.addEventListener("pointerdown", (event) => {
      select.dataset.singleChange = event.altKey ? "true" : "false";
    });
    select.addEventListener("keydown", (event) => {
      select.dataset.singleChange = event.altKey ? "true" : "false";
    });
    select.addEventListener("click", (event) => event.stopPropagation());
    select.addEventListener("change", (event) => {
      onChooseReplacementIdentifier(mark, event.target.value, event.target.dataset.singleChange === "true");
      event.target.dataset.singleChange = "false";
    });
    wrap.appendChild(select);
    return wrap;
  }

  const button = document.createElement("button");
  button.className = `review-highlight ${mark.kind}`;
  button.type = "button";
  button.textContent = content;
  button.title = `Pseudonymise: ${mark.reason}. Shift-click to remove all matching text.`;
  button.addEventListener("click", (event) => {
    event.stopPropagation();
    if (event.shiftKey) {
      onRemoveSuggestion(mark);
      return;
    }
    onApplySuggestion(mark);
  });
  return button;
}

function onRemoveSuggestion(mark) {
  setReviewText(reviewTextValue.split(mark.value).join(""));
  setStatus(`Removed all matching ${mark.kind === "likely" ? "direct" : "possible"} identifiers.`);
}

function onApplySuggestion(mark) {
  const placeholder = mark.replacement || nextManualPlaceholder(mark.category, buildEffectiveMapping());
  manualMapping = { ...manualMapping, [placeholder]: mark.value };
  setReviewText(reviewTextValue.split(mark.value).join(placeholder));
  setStatus(`Replaced all matching ${mark.kind === "likely" ? "direct" : "possible"} identifiers with ${placeholder}.`);
}

function onChooseReplacementIdentifier(mark, value, singleInstance = false) {
  const placeholder = placeholderFromIdentifierOption(value, buildEffectiveMapping());
  if (!placeholder || placeholder === mark.placeholder) {
    return;
  }

  manualMapping = { ...manualMapping, [placeholder]: mark.original };
  if (singleInstance) {
    replaceReviewRange(mark.start, mark.end, placeholder);
    setStatus(`Changed one ${mark.placeholder} to ${placeholder}.`);
    return;
  }

  setReviewText(reviewTextValue.split(mark.placeholder).join(placeholder));
  setStatus(`Changed all ${mark.placeholder} instances to ${placeholder}.`);
}

function buildReviewMarks(text, originalPseudonymisedText, mapping, residualSpans) {
  if (!text) {
    return [];
  }

  const replacedMarks = Object.entries(mapping).flatMap(([placeholder, original]) =>
    findAllOccurrences(text, placeholder).map((start) => ({
      kind: "replaced",
      start,
      end: start + placeholder.length,
      placeholder,
      original,
    }))
  );
  const revertedMarks = Object.entries(mapping).flatMap(([placeholder, original]) =>
    findAllOccurrences(text, original).map((start) => ({
      kind: "likely",
      start,
      end: start + original.length,
      category: placeholderCategoryValue(placeholder),
      reason: "Previously pseudonymised value was restored",
      value: original,
      replacement: placeholder,
    }))
  );
  const suggestionMarks = residualSpans.flatMap((span) => {
    const value = originalPseudonymisedText.slice(span.start, span.end);
    if (!value || mapping[value]) {
      return [];
    }

    return findAllOccurrences(text, value).map((start) => ({
      kind: span.severity === "direct" ? "likely" : "possible",
      start,
      end: start + value.length,
      category: span.category,
      reason: span.reason,
      value,
    }));
  });

  return normaliseReviewMarks([...replacedMarks, ...revertedMarks, ...suggestionMarks], text.length);
}

function findAllOccurrences(text, value) {
  const starts = [];
  let cursor = 0;
  while (value && cursor < text.length) {
    const start = text.indexOf(value, cursor);
    if (start === -1) {
      break;
    }
    starts.push(start);
    cursor = start + value.length;
  }
  return starts;
}

function normaliseReviewMarks(marks, textLength) {
  const ordered = marks
    .filter((mark) => mark.start >= 0 && mark.end <= textLength && mark.start < mark.end)
    .sort(
      (left, right) =>
        left.start - right.start ||
        reviewMarkRank(left.kind) - reviewMarkRank(right.kind) ||
        right.end - left.end
    );
  const accepted = [];
  let cursor = 0;

  for (const mark of ordered) {
    if (mark.start >= cursor) {
      accepted.push(mark);
      cursor = mark.end;
    }
  }

  return accepted;
}

function reviewMarkRank(kind) {
  if (kind === "replaced") {
    return 0;
  }
  return kind === "likely" ? 1 : 2;
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

function buildIdentifierOptions(mapping) {
  const existing = Object.keys(mapping || {}).filter(
    (placeholder) =>
      /^\[[A-Z_]+_\d{3}\]$/.test(placeholder) &&
      !placeholder.includes("_PREFERRED_NAME_") &&
      !placeholder.includes("_FIRST_NAME_") &&
      !placeholder.includes("_SURNAME_")
  );
  const newOptions = PLACEHOLDER_CATEGORY_OPTIONS.map((category) => `${category}_NEW`);
  return [...existing, ...newOptions].sort((left, right) =>
    identifierOptionLabel(left).localeCompare(identifierOptionLabel(right))
  );
}

function placeholderFromIdentifierOption(value, mapping) {
  if (String(value || "").endsWith("_NEW")) {
    return nextManualPlaceholder(String(value).slice(0, -4), mapping);
  }
  return String(value || "").startsWith("[") ? String(value) : null;
}

function identifierOptionLabel(value) {
  return String(value || "").startsWith("[") ? String(value).slice(1, -1) : String(value || "");
}

function placeholderCategoryValue(placeholder) {
  const match = String(placeholder || "").match(/^\[([A-Z_]+)_\d{3}\]$/);
  return match?.[1] || "IDENTIFIER";
}

function normaliseResidualSpans(spans, textLength) {
  const ordered = (Array.isArray(spans) ? spans : [])
    .filter((span) => span.start >= 0 && span.end <= textLength && span.start < span.end)
    .sort(
      (left, right) =>
        left.start - right.start ||
        severityRank(left.severity) - severityRank(right.severity) ||
        right.end - left.end
    );
  const accepted = [];
  let cursor = 0;

  for (const span of ordered) {
    if (span.start >= cursor) {
      accepted.push(span);
      cursor = span.end;
    }
  }

  return accepted;
}

function severityRank(severity) {
  return severity === "direct" ? 0 : 1;
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
  updateExportControls();
  if (!placeholderText.value.trim()) {
    restoreStatus.textContent = "Paste placeholder-bearing output to restore with the local map.";
    return;
  }
  restoreStatus.textContent = `${restored.restoredCount.toLocaleString()} restored${
    restored.unresolvedCount > 0 ? `, ${restored.unresolvedCount.toLocaleString()} unresolved` : ""
  }.`;
}

function buildExportTitle() {
  const mode = REPORT_MODES[getGeneratedReportType()] || REPORT_MODES.wellbeing_assurance_visit;
  return mode.label || "Client data report";
}

function buildReportHtml(text) {
  const paragraphs = String(text || "")
    .split(/\n{2,}/)
    .map((block) => block.trim())
    .filter(Boolean)
    .map((block) => `<p>${escapeHtml(block).replace(/\n/g, "<br>")}</p>`)
    .join("\n");
  return [
    "<article>",
    `<h1>${escapeHtml(buildExportTitle())}</h1>`,
    paragraphs || "<p></p>",
    "</article>",
  ].join("\n");
}

function safeFilenamePart(value, fallback) {
  return String(value || fallback || "")
    .trim()
    .replace(/[\\/:*?"<>|]+/g, "")
    .replace(/\s+/g, " ")
    .slice(0, 80) || fallback;
}

function parseReportDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value;
  }
  const raw = String(value || "").trim();
  if (!raw) {
    return null;
  }

  const isoMatch = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (isoMatch) {
    const [, year, month, day] = isoMatch;
    return new Date(Number(year), Number(month) - 1, Number(day));
  }

  const ukMatch = raw.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/);
  if (ukMatch) {
    const [, day, month, year] = ukMatch;
    const fullYear = year.length === 2 ? Number(`20${year}`) : Number(year);
    return new Date(fullYear, Number(month) - 1, Number(day));
  }

  const parsed = new Date(raw);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function formatReportDate(value, fallbackDate = new Date()) {
  const date = parseReportDate(value) || fallbackDate;
  return new Intl.DateTimeFormat("en-GB", {
    day: "numeric",
    month: "long",
    year: "numeric",
  }).format(date);
}

function buildReportDownloadFilename(report, extension, fallbackDate = new Date()) {
  const clientDetails = report?.client_details && typeof report.client_details === "object" ? report.client_details : {};
  const clientName =
    clientDetails.client_name ||
    clientDetails.name ||
    clientDetails.full_name ||
    clientDetails.preferred_name ||
    "Client";
  const reportDate =
    clientDetails.report_date ||
    clientDetails.visit_date ||
    clientDetails.assessment_date ||
    report?.report_date ||
    report?.date;
  const title = report?.report_title || buildExportTitle();
  return [
    safeFilenamePart(clientName, "Client"),
    safeFilenamePart(formatReportDate(reportDate, fallbackDate), "Date"),
    safeFilenamePart(title, "Report"),
  ].join(" - ") + `.${extension}`;
}

function restoreReportPlaceholders(value, mapping = buildEffectiveMapping()) {
  if (typeof value === "string") {
    return deanonymiseText(mapping, value).text;
  }
  if (Array.isArray(value)) {
    return value.map((item) => restoreReportPlaceholders(item, mapping));
  }
  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value).map(([key, item]) => [key, restoreReportPlaceholders(item, mapping)])
    );
  }
  return value;
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function normalizePdfText(value) {
  return String(value || "")
    .replace(/[“”]/g, '"')
    .replace(/[‘’]/g, "'")
    .replace(/[–—]/g, "-")
    .replace(/…/g, "...")
    .replace(/[^\x09\x0A\x0D\x20-\x7E]/g, "");
}

function escapePdfText(value) {
  return normalizePdfText(value).replace(/\\/g, "\\\\").replace(/\(/g, "\\(").replace(/\)/g, "\\)");
}

function wrapPdfLine(text, maxChars) {
  const words = normalizePdfText(text).split(/\s+/).filter(Boolean);
  const lines = [];
  let line = "";

  for (const word of words) {
    if (!line) {
      line = word;
      continue;
    }
    if (`${line} ${word}`.length <= maxChars) {
      line = `${line} ${word}`;
      continue;
    }
    lines.push(line);
    line = word;
  }

  if (line) {
    lines.push(line);
  }
  return lines.length ? lines : [""];
}

function buildPdfBlob({ title, text }) {
  const pageWidth = 595;
  const pageHeight = 842;
  const margin = 54;
  const titleSize = 18;
  const bodySize = 11;
  const lineHeight = 15;
  const maxLinesPerPage = Math.floor((pageHeight - margin * 2 - titleSize - 18) / lineHeight);
  const bodyLines = [];

  for (const block of normalizePdfText(text).split(/\n{2,}/)) {
    const trimmed = block.trim();
    if (!trimmed) {
      continue;
    }
    for (const line of trimmed.split(/\n/)) {
      bodyLines.push(...wrapPdfLine(line, 88));
    }
    bodyLines.push("");
  }
  if (bodyLines.at(-1) === "") {
    bodyLines.pop();
  }

  const pages = [];
  for (let index = 0; index < bodyLines.length || index === 0; index += maxLinesPerPage) {
    pages.push(bodyLines.slice(index, index + maxLinesPerPage));
  }

  const objects = [];
  const addObject = (content) => {
    objects.push(content);
    return objects.length;
  };

  const catalogId = addObject("<< /Type /Catalog /Pages 2 0 R >>");
  const pagesId = addObject("");
  const fontId = addObject("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>");
  const pageIds = [];

  pages.forEach((lines) => {
    const commands = [
      "BT",
      `/F1 ${titleSize} Tf`,
      `${margin} ${pageHeight - margin} Td`,
      `(${escapePdfText(title)}) Tj`,
      `/F1 ${bodySize} Tf`,
      `0 -${titleSize + 18} Td`,
      `${lineHeight} TL`,
    ];

    lines.forEach((line, index) => {
      if (index > 0) {
        commands.push("T*");
      }
      commands.push(`(${escapePdfText(line)}) Tj`);
    });
    commands.push("ET");

    const stream = commands.join("\n");
    const contentId = addObject(`<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`);
    const pageId = addObject(
      `<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 ${pageWidth} ${pageHeight}] /Resources << /Font << /F1 ${fontId} 0 R >> >> /Contents ${contentId} 0 R >>`
    );
    pageIds.push(pageId);
  });

  objects[pagesId - 1] = `<< /Type /Pages /Kids [${pageIds.map((id) => `${id} 0 R`).join(" ")}] /Count ${pageIds.length} >>`;

  let pdf = "%PDF-1.4\n";
  const offsets = [0];
  objects.forEach((content, index) => {
    offsets.push(pdf.length);
    pdf += `${index + 1} 0 obj\n${content}\nendobj\n`;
  });
  const xrefOffset = pdf.length;
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (let index = 1; index < offsets.length; index += 1) {
    pdf += `${String(offsets[index]).padStart(10, "0")} 00000 n \n`;
  }
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root ${catalogId} 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`;

  return new Blob([pdf], { type: "application/pdf" });
}

function stripLeadingPdfTitle(text, title) {
  const value = String(text || "").trim();
  const heading = String(title || "").trim();
  if (!heading || !value.startsWith(heading)) {
    return value;
  }
  return value.slice(heading.length).replace(/^\s+/, "");
}

function exportPdf() {
  const text = restoredText.value.trim();
  if (!text) {
    setExportStatus("Restore text before exporting.", true);
    updateExportControls();
    return;
  }

  const reportType = getGeneratedReportType();
  if (reportType !== "simple_summary" && structuredReport?.status === "ready_for_render") {
    const restoredReport = restoreReportPlaceholders(structuredReport);
    const title = restoredReport.report_title || buildExportTitle();
    const reportText = stripLeadingPdfTitle(buildStructuredReportText(restoredReport), title);
    downloadBlob(
      buildPdfBlob({ title, text: reportText }),
      buildReportDownloadFilename(restoredReport, "pdf")
    );
    setExportStatus("PDF exported.");
    return;
  }

  downloadBlob(buildPdfBlob({ title: buildExportTitle(), text }), "client-data-report.pdf");
  setExportStatus("PDF exported.");
}

async function exportWordDoc() {
  const text = restoredText.value.trim();
  if (!text) {
    setExportStatus("Restore text before exporting.", true);
    updateExportControls();
    return;
  }
  const reportType = getGeneratedReportType();
  if (reportType === "simple_summary") {
    const html = [
      "<!doctype html>",
      "<html>",
      "<head>",
      '<meta charset="utf-8">',
      `<title>${escapeHtml(buildExportTitle())}</title>`,
      "<style>",
      "body{font-family:Arial,sans-serif;line-height:1.45;color:#1c2533;}",
      "h1{font-size:20pt;margin:0 0 16pt;}",
      "p{font-size:11pt;margin:0 0 10pt;}",
      "</style>",
      "</head>",
      "<body>",
      buildReportHtml(text),
      "</body>",
      "</html>",
    ].join("\n");
    const blob = new Blob([html], { type: "application/msword;charset=utf-8" });
    downloadBlob(blob, "client-data-report.doc");
    setExportStatus("Word document exported.");
    return;
  }
  if (!structuredReport || structuredReport.status !== "ready_for_render") {
    setExportStatus("Generate a structured report before exporting Word.", true);
    updateExportControls();
    return;
  }

  try {
    setExportStatus("Generating Word document...");
    const restoredReport = restoreReportPlaceholders(structuredReport);
    const exportResult = await directoryApi.exportStructuredReportDocx({
      reportType,
      report: restoredReport,
    });
    downloadBlob(exportResult.blob || exportResult, exportResult.filename || "wellbeing-assurance-visit-summary.docx");
    setExportStatus("Word document exported.");
  } catch (error) {
    setExportStatus(error?.message || "Could not export Word document.", true);
  }
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

  const marks = getReviewMarks();
  const reviewMarks = marks.filter((mark) => mark.kind !== "replaced");
  summaryMessage.textContent =
    reviewMarks.length > 0 ? "Needs review: resolve orange and yellow identifiers before copying." : "Ready: only replacement tags remain.";

  if (!reviewMarks.length) {
    residualList.innerHTML = '<p class="muted">Ready. No visible residual identifiers.</p>';
  } else {
    for (const mark of reviewMarks.slice(0, 12)) {
      const row = document.createElement("div");
      row.className = "client-data-list-row";
      row.innerHTML = `
        <div>
          <strong>${escapeHtml(mark.value)}</strong>
          <span>${escapeHtml(mark.category || "IDENTIFIER")} · ${escapeHtml(mark.reason || "Residual identifier")} · ${mark.kind === "likely" ? "needs review" : "possible"}</span>
        </div>
      `;
      const button = document.createElement("button");
      button.className = "secondary";
      button.type = "button";
      button.textContent = "Replace";
      button.addEventListener("click", () => onApplySuggestion(mark));
      row.appendChild(button);
      residualList.appendChild(row);
    }
    if (reviewMarks.length > 12) {
      const more = document.createElement("p");
      more.className = "muted";
      more.textContent = `${(reviewMarks.length - 12).toLocaleString()} more review item${reviewMarks.length - 12 === 1 ? "" : "s"} visible in the highlighted text.`;
      residualList.appendChild(more);
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
placeholderText.addEventListener("input", updateRestore);
preferredNameInput.addEventListener("input", () => {
  renderReviewOutput();
  renderSummary();
  updateRestore();
});
reviewOutput.addEventListener("click", (event) => {
  if (!event.shiftKey || event.target.closest("button, select, input, textarea")) {
    return;
  }
  if (onApplySelectedReviewText()) {
    event.preventDefault();
  }
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
exportPdfBtn?.addEventListener("click", exportPdf);
exportWordBtn?.addEventListener("click", () => {
  void exportWordDoc();
});
generateReportBtn?.addEventListener("click", () => {
  void generateReport();
});
reportRevisionText?.addEventListener("input", updateRevisionControls);
reviseReportBtn?.addEventListener("click", () => {
  void reviseReport();
});
reportModelLockBtn?.addEventListener("click", () => {
  const mode = getSelectedReportMode();
  if (!mode.modelLocked) {
    return;
  }
  reportModelUnlocked = !reportModelUnlocked;
  updateReportControls();
  setReportStatus(
    reportModelUnlocked
      ? "Model controls unlocked for this report."
      : `${mode.label} model controls locked to the default.`
  );
});
reportProviderSelect?.addEventListener("change", () => {
  populateReportModels();
  syncReportThinkingControl();
  setReportStatus(`${reportProviderSelect.options[reportProviderSelect.selectedIndex]?.text || "Provider"} selected.`);
});
reportModelSelect?.addEventListener("change", () => {
  setReportStatus(`${reportModelSelect.options[reportModelSelect.selectedIndex]?.text || "Model"} selected.`);
});
reportThinkingSelect?.addEventListener("change", () => {
  setReportStatus(`${reportThinkingSelect.options[reportThinkingSelect.selectedIndex]?.text || "Thinking"} thinking selected.`);
});
reportModeSelect?.addEventListener("change", () => {
  const mode = getSelectedReportMode();
  applyReportModeDefaults();
  setReportStatus(`${mode.label} selected.`);
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
renderReviewOutput();
applyReportModeDefaults();
void init();
