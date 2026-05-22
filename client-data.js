import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260512";
import {
  DEFAULT_GAP_PX,
  DEFAULT_RADIUS_PX,
  EXPORT_WIDTH,
  LAYOUTS,
  LAYOUT_BACKGROUND_COLORS,
  MAX_SELECTION,
  clamp,
  computeAdjustedSlotRect,
  computeImagePlacement,
  drawImageIntoSlot,
  drawRoundedRectPath,
  getSlotNeighborFlags,
} from "./photo-layout-core.js";

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

const CLIENT_LIST_CACHE_KEY = "clientDataReportPhotoClientsV1";
const CLIENT_LIST_CACHE_TTL_MS = 6 * 60 * 60 * 1000;

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
const reportDraftPanel = document.getElementById("reportDraftPanel");
const showPlaceholderText = document.getElementById("showPlaceholderText");
const reportNotesPanel = document.getElementById("reportNotesPanel");
const reportProviderSelect = document.getElementById("reportProviderSelect");
const reportModelSelect = document.getElementById("reportModelSelect");
const reportThinkingField = document.getElementById("reportThinkingField");
const reportThinkingSelect = document.getElementById("reportThinkingSelect");
const reportModeSelect = document.getElementById("reportModeSelect");
const reportModelLockBtn = document.getElementById("reportModelLockBtn");
const reportGuidanceText = document.getElementById("reportGuidanceText");
const reportAudienceSelect = document.getElementById("reportAudienceSelect");
const reportStyleSelect = document.getElementById("reportStyleSelect");
const generateReportBtn = document.getElementById("generateReportBtn");
const reportRevisionPanel = document.getElementById("reportRevisionPanel");
const reportRevisionText = document.getElementById("reportRevisionText");
const reportRevisionHistory = document.getElementById("reportRevisionHistory");
const reviseReportBtn = document.getElementById("reviseReportBtn");
const reportPhotoClientSelect = document.getElementById("reportPhotoClientSelect");
const reportPhotoLoadBtn = document.getElementById("reportPhotoLoadBtn");
const reportPhotoStatus = document.getElementById("reportPhotoStatus");
const reportPhotoImagesGrid = document.getElementById("reportPhotoImagesGrid");
const reportPhotoLayoutPicker = document.getElementById("reportPhotoLayoutPicker");
const reportPhotoStage = document.getElementById("reportPhotoStage");
const reportPhotoSelectedList = document.getElementById("reportPhotoSelectedList");
const reportPhotoLocalDropZone = document.getElementById("reportPhotoLocalDropZone");
const reportPhotoLocalInput = document.getElementById("reportPhotoLocalInput");
const reportPhotoZoomRange = document.getElementById("reportPhotoZoomRange");
const reportPhotoPanXRange = document.getElementById("reportPhotoPanXRange");
const reportPhotoPanYRange = document.getElementById("reportPhotoPanYRange");
const reportPhotoResetBtn = document.getElementById("reportPhotoResetBtn");
const reportPhotoGapEnabled = document.getElementById("reportPhotoGapEnabled");
const reportPhotoGapRange = document.getElementById("reportPhotoGapRange");
const reportPhotoGapValue = document.getElementById("reportPhotoGapValue");
const reportPhotoRoundedEnabled = document.getElementById("reportPhotoRoundedEnabled");
const reportPhotoCornerRange = document.getElementById("reportPhotoCornerRange");
const reportPhotoCornerValue = document.getElementById("reportPhotoCornerValue");
const reportPhotoBackgroundPicker = document.getElementById("reportPhotoBackgroundPicker");
const generateCollageBtn = document.getElementById("generateCollageBtn");
const clearCollageBtn = document.getElementById("clearCollageBtn");
const reportCollagePreview = document.getElementById("reportCollagePreview");
const exportStatus = document.getElementById("exportStatus");
const exportPdfBtn = document.getElementById("exportPdfBtn");
const exportWordBtn = document.getElementById("exportWordBtn");

let account = null;
let result = null;
let reviewTextValue = "";
let manualMapping = {};
let ignoredReviewValues = new Set();
let busy = false;
let reportBusy = false;
let structuredReport = null;
let reportSourceNotes = "";
let reportModelUnlocked = false;
let reportRevisionRequests = [];
let reportPhotoClients = [];
let reportPhotoPool = [];
let reportPhotoSelected = [];
let reportPhotoSelectedSlot = -1;
let reportPhotoActiveLayoutId = LAYOUTS[0].id;
let reportPhotoDragState = null;
let reportPhotoCollage = null;
let reportPhotoCollageUrl = "";
const reportPhotoImageMetaCache = new Map();
const reportPhotoExportImageCache = new Map();
const reportPhotoStyle = {
  backgroundColor: "#ffffff",
  gapEnabled: true,
  gapPx: DEFAULT_GAP_PX,
  roundedEnabled: true,
  cornerRadiusPx: DEFAULT_RADIUS_PX,
};

const LOW_SIGNAL_PRONOUNS = new Set(["he", "him", "his", "she", "her", "hers"]);
const PERSON_PLACEHOLDER_CATEGORIES = new Set(["CLIENT", "STAFF", "RELATIVE", "FRIEND", "PROFESSIONAL"]);

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

function setDraftComplete(message = "Changes complete.") {
  if (reportDraftPanel) {
    reportDraftPanel.classList.add("is-complete");
  }
  if (restoreStatus) {
    restoreStatus.textContent = message;
  }
}

function clearDraftComplete() {
  if (reportDraftPanel) {
    reportDraftPanel.classList.remove("is-complete");
  }
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
  if (reportGuidanceText) {
    reportGuidanceText.value = "";
  }
  if (reportAudienceSelect) {
    reportAudienceSelect.value = "Professional";
  }
  if (reportStyleSelect) {
    reportStyleSelect.value = "Professional";
  }
  result = null;
  manualMapping = {};
  ignoredReviewValues = new Set();
  structuredReport = null;
  reportSourceNotes = "";
  reportRevisionRequests = [];
  reportPhotoSelected.forEach((image) => {
    if (image?.localObjectUrl) {
      URL.revokeObjectURL(image.localObjectUrl);
    }
  });
  reportPhotoSelected = [];
  reportPhotoPool = [];
  reportPhotoSelectedSlot = -1;
  invalidateReportCollage();
  clearDraftComplete();
  if (reportDraftPanel) {
    reportDraftPanel.hidden = true;
  }
  if (showPlaceholderText) {
    showPlaceholderText.checked = false;
  }
  if (reportRevisionText) {
    reportRevisionText.value = "";
  }
  renderReportNotes(null);
  renderRevisionHistory();
  syncDraftTextVisibility();
  if (summaryDetails) {
    delete summaryDetails.dataset.userOpened;
    summaryDetails.open = false;
  }
  setRiskBadge(null);
  setStatus("Paste a note to begin.");
  setReportStatus("Generate a placeholder-preserving report from the pseudonymised text.");
  setReportPhotoStatus("Optional: create a collage for the report image section.");
  setExportStatus("Export or copy the draft report text.");
  renderReviewOutput();
  renderReportPhotoComposer();
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
    ignoredReviewValues = new Set();
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
  setExportStatus("Copied draft report text.");
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
  return parts.filter(Boolean).join("\n\n");
}

function renderReportNotes(report) {
  if (!reportNotesPanel) {
    return;
  }
  const groups = [
    ["Warnings", report?.warnings],
    ["Clarification Notes", report?.clarification_notes],
    ["Assumptions Avoided", report?.assumptions_avoided],
  ]
    .map(([title, items]) => [title, (Array.isArray(items) ? items : []).map(stringifyReportValue).filter(Boolean)])
    .filter(([, items]) => items.length > 0);

  reportNotesPanel.hidden = groups.length === 0;
  reportNotesPanel.innerHTML = groups
    .map(
      ([title, items]) => `
        <div class="client-data-report-note-group">
          <h4>${escapeHtml(title)}</h4>
          <ul>${items.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>
        </div>
      `
    )
    .join("");
}

function renderRevisionHistory() {
  if (!reportRevisionHistory) {
    return;
  }
  reportRevisionHistory.hidden = reportRevisionRequests.length === 0;
  reportRevisionHistory.innerHTML = reportRevisionRequests.length
    ? `
        <h4>Requested changes</h4>
        <ol>${reportRevisionRequests.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ol>
      `
    : "";
}

function syncDraftTextVisibility() {
  const showPlaceholders = Boolean(showPlaceholderText?.checked);
  if (placeholderText) {
    placeholderText.hidden = !showPlaceholders;
  }
  if (restoredText) {
    restoredText.hidden = showPlaceholders;
  }
}

function setStructuredReport(report) {
  structuredReport = report && typeof report === "object" ? report : null;
  if (!structuredReport) {
    clearDraftComplete();
    placeholderText.value = "";
    updateRestore();
    if (reportDraftPanel) {
      reportDraftPanel.hidden = true;
    }
    renderReportNotes(null);
    updateRevisionControls();
    updateExportControls();
    return;
  }

  if (reportDraftPanel) {
    reportDraftPanel.hidden = false;
  }
  clearDraftComplete();
  if (structuredReport.status === "needs_notes") {
    placeholderText.value = "";
    updateRestore();
    renderReportNotes(structuredReport);
    setReportStatus("Please provide notes.", true);
  } else if (structuredReport.status === "needs_clarification") {
    placeholderText.value = buildStructuredReportText(structuredReport);
    updateRestore();
    renderReportNotes(structuredReport);
    setReportStatus("Clarification needed before this report can be rendered.", true);
  } else {
    placeholderText.value = buildStructuredReportText(structuredReport);
    updateRestore();
    renderReportNotes(structuredReport);
    setReportStatus("Draft report generated. Review it below, request changes, or export when ready.");
    setDraftComplete("Draft ready.");
  }
  syncDraftTextVisibility();
  updateRevisionControls();
  updateExportControls();
  renderRevisionHistory();
}

async function generateReport(revisionRequest = "") {
  if (reportBusy) {
    return;
  }

  setReportBusy(true);
  const thinkingOptions = getSelectedThinkingOptions();
  const isRevision = Boolean(String(revisionRequest || "").trim() && structuredReport);
  clearDraftComplete();
  setReportStatus(isRevision ? "Regenerating report..." : "Generating report...");
  try {
    const response = await directoryApi.generateStructuredReport({
      reportType: getSelectedReportType(),
      notes: isRevision ? reportSourceNotes || reviewTextValue : reviewTextValue,
      provider: getSelectedReportProvider(),
      model: reportModelSelect?.value,
      thinking: thinkingOptions.thinking,
      reasoningEffort: thinkingOptions.reasoningEffort,
      guidance: reportGuidanceText?.value.trim() || "",
      audience: reportAudienceSelect?.value || "Professional",
      writingStyle: reportStyleSelect?.value || "Professional",
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
    if (isRevision) {
      reportRevisionRequests.push(normalisePreferredName(revisionRequest) || String(revisionRequest || "").trim());
      renderRevisionHistory();
      if (reportRevisionText) {
        reportRevisionText.value = "";
      }
      setDraftComplete("Changes complete.");
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
  if (isLowSignalPronoun(value)) {
    setStatus("Pronouns are not treated as identifiers.");
    updateManualControls();
    return;
  }

  const mapping = buildEffectiveMapping();
  const placeholder = nextManualPlaceholder(manualIdentifierCategory.value, mapping);
  manualMapping = { ...manualMapping, [placeholder]: value };
  setReviewText(replaceAllOccurrences(reviewTextValue, value, placeholder));
  manualIdentifierText.value = "";
  setStatus(`Added ${placeholder}.`);
}

function applyResidual(span) {
  const value = String(result?.pseudonymised_text || "").slice(span.start, span.end);
  if (!value || !reviewTextValue.includes(value)) {
    return;
  }
  if (isLowSignalPronoun(value)) {
    setStatus("Pronouns are not treated as identifiers.");
    return;
  }

  const mapping = buildEffectiveMapping();
  const placeholder = nextManualPlaceholder(span.category || "IDENTIFIER", mapping);
  manualMapping = { ...manualMapping, [placeholder]: value };
  setReviewText(replaceAllOccurrences(reviewTextValue, value, placeholder, { caseInsensitive: true }));
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
  if (isLowSignalPronoun(value)) {
    setStatus("Pronouns are not treated as identifiers.");
    window.getSelection?.().removeAllRanges();
    return true;
  }
  if (/^\[[A-Za-z_]+_\d{3}\]$/.test(value)) {
    setStatus("Selected text is already a placeholder.");
    window.getSelection?.().removeAllRanges();
    return true;
  }

  const mapping = buildEffectiveMapping();
  const placeholder = nextManualPlaceholder(manualIdentifierCategory.value || "IDENTIFIER", mapping);
  manualMapping = { ...manualMapping, [placeholder]: value };
  setReviewText(replaceAllOccurrences(reviewTextValue, value, placeholder, { caseInsensitive: true }));
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
  button.title = `Pseudonymise: ${mark.reason}. Shift-click to remove this highlight for all matching text.`;
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
  if (!mark.value) {
    return;
  }
  ignoredReviewValues.add(normaliseIgnoredReviewValue(mark.value));
  renderReviewOutput();
  renderSummary();
  setStatus(`Removed highlighting for all matching ${mark.kind === "likely" ? "direct" : "possible"} identifiers.`);
}

function onApplySuggestion(mark) {
  const placeholder = mark.replacement || nextManualPlaceholder(mark.category, buildEffectiveMapping());
  ignoredReviewValues.delete(normaliseIgnoredReviewValue(mark.value));
  manualMapping = { ...manualMapping, [placeholder]: mark.value };
  setReviewText(replaceAllOccurrences(reviewTextValue, mark.value, placeholder, { caseInsensitive: mark.caseInsensitive }));
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

  setReviewText(replaceAllOccurrences(reviewTextValue, mark.placeholder, placeholder, { caseInsensitive: true }));
  setStatus(`Changed all ${mark.placeholder} instances to ${placeholder}.`);
}

function buildReviewMarks(text, originalPseudonymisedText, mapping, residualSpans) {
  if (!text) {
    return [];
  }

  const replacedMarks = Object.entries(mapping).flatMap(([placeholder, original]) =>
    findAllOccurrences(text, placeholder, { caseInsensitive: true }).map((start) => ({
      kind: "replaced",
      start,
      end: start + placeholder.length,
      placeholder,
      original,
    }))
  );
  const revertedMarks = Object.entries(mapping).flatMap(([placeholder, original]) => {
    const info = parseBasePlaceholder(placeholder);
    if (!info || isLowSignalPronoun(original)) {
      return [];
    }

    const marks = restoredValueMarks(placeholder, original, info);
    return marks.flatMap((mark) =>
      findAllOccurrences(text, mark.value, { caseInsensitive: true }).map((start) => ({
        kind: "likely",
        start,
        end: start + mark.value.length,
        category: info.category,
        reason: "Previously pseudonymised value was restored",
        value: mark.value,
        replacement: mark.replacement,
        caseInsensitive: true,
      }))
    );
  });
  const suggestionMarks = residualSpans.flatMap((span) => {
    const value = originalPseudonymisedText.slice(span.start, span.end);
    if (!value || mapping[value] || isLowSignalPronoun(value) || ignoredReviewValues.has(normaliseIgnoredReviewValue(value))) {
      return [];
    }

    return findAllOccurrences(text, value, { caseInsensitive: true }).map((start) => ({
      kind: span.severity === "direct" ? "likely" : "possible",
      start,
      end: start + value.length,
      category: span.category,
      reason: span.reason,
      value,
      caseInsensitive: true,
    }));
  });

  return normaliseReviewMarks(
    [
      ...replacedMarks,
      ...revertedMarks.filter((mark) => !ignoredReviewValues.has(normaliseIgnoredReviewValue(mark.value))),
      ...suggestionMarks,
    ],
    text.length
  );
}

function findAllOccurrences(text, value, options = {}) {
  const starts = [];
  const needle = String(value || "");
  if (!needle) {
    return starts;
  }
  const haystack = options.caseInsensitive ? String(text || "").toLowerCase() : String(text || "");
  const searchNeedle = options.caseInsensitive ? needle.toLowerCase() : needle;
  let cursor = 0;
  while (cursor < haystack.length) {
    const start = haystack.indexOf(searchNeedle, cursor);
    if (start === -1) {
      break;
    }
    starts.push(start);
    cursor = start + needle.length;
  }
  return starts;
}

function restoredValueMarks(placeholder, original, info) {
  const baseMark = [{ value: original, replacement: placeholder }];
  if (!PERSON_PLACEHOLDER_CATEGORIES.has(info.category)) {
    return baseMark;
  }

  const { firstName, surname } = splitPersonName(original);
  const personMarks = [...baseMark];
  if (firstName && !isLowSignalPronoun(firstName)) {
    personMarks.push({
      value: firstName,
      replacement: `[${info.category}_PREFERRED_NAME_${info.index}]`,
    });
  }
  if (surname && !isLowSignalPronoun(surname)) {
    personMarks.push({
      value: surname,
      replacement: `[${info.category}_SURNAME_${info.index}]`,
    });
  }

  return dedupeValueMarks(personMarks);
}

function dedupeValueMarks(marks) {
  const seen = new Set();
  return marks.filter((mark) => {
    const key = `${normaliseIgnoredReviewValue(mark.value)}|${mark.replacement}`;
    if (!mark.value || seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

function replaceAllOccurrences(text, value, replacement, options = {}) {
  return findAllOccurrences(text, value, options)
    .reverse()
    .reduce(
      (nextText, start) => `${nextText.slice(0, start)}${replacement}${nextText.slice(start + String(value).length)}`,
      String(text || "")
    );
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
  const merged = { ...(result?.mapping || {}) };
  for (const [placeholder, value] of Object.entries(manualMapping)) {
    merged[normalisePlaceholderKey(placeholder)] = value;
  }
  const preferredClientName = normalisePreferredName(preferredNameInput.value);

  for (const [placeholder, original] of Object.entries({ ...merged })) {
    const match = normalisePlaceholderKey(placeholder).match(/^\[([A-Z_]+)_(\d{3})\]$/);
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
  return parseBasePlaceholder(placeholder)?.category || "IDENTIFIER";
}

function parseBasePlaceholder(placeholder) {
  const match = normalisePlaceholderKey(placeholder).match(/^\[([A-Z_]+)_(\d{3})\]$/);
  if (!match) {
    return null;
  }
  return { category: match[1], index: match[2] };
}

function normalisePlaceholderKey(placeholder) {
  const match = String(placeholder || "").match(/^([\[<])([A-Za-z0-9_]+)([\]>])$/);
  if (!match) {
    return String(placeholder || "");
  }
  const open = match[1] === "<" ? "<" : "[";
  const close = open === "<" ? ">" : "]";
  return `${open}${match[2].toUpperCase()}${close}`;
}

function isLowSignalPronoun(value) {
  return LOW_SIGNAL_PRONOUNS.has(String(value || "").trim().toLowerCase());
}

function normaliseIgnoredReviewValue(value) {
  return String(value || "").trim().toLowerCase();
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
  const output = String(text || "").replace(/(?:\[[A-Za-z0-9_]+\]|<[A-Za-z0-9_]+>)/g, (placeholder) => {
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
  const normalisedPlaceholder = normalisePlaceholderKey(placeholder);
  if (mapping[placeholder]) {
    return mapping[placeholder];
  }
  if (mapping[normalisedPlaceholder]) {
    return mapping[normalisedPlaceholder];
  }

  const match = normalisedPlaceholder.match(/^[<\[]([A-Z_]+)_(PREFERRED_NAME|FIRST_NAME|SURNAME)_(\d{3})[>\]]$/);
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
    restoreStatus.textContent = "Generated report text will appear here.";
    return;
  }
  restoreStatus.textContent =
    restored.unresolvedCount > 0
      ? `Draft ready with ${restored.unresolvedCount.toLocaleString()} unresolved placeholder${restored.unresolvedCount === 1 ? "" : "s"}.`
      : "Draft ready.";
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

function escapedRegExp(value) {
  return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function titleClientName(title) {
  const match = String(title || "").match(/\bfor\s+(.+)$/i);
  return match ? match[1].trim() : "";
}

function cleanReportTitle(title, clientName = "") {
  const raw = String(title || buildExportTitle()).trim();
  if (!clientName) {
    return raw;
  }
  return raw.replace(new RegExp(`\\s+for\\s+${escapedRegExp(clientName)}\\s*$`, "i"), "").trim() || raw;
}

function buildReportDownloadFilename(report, extension, fallbackDate = new Date()) {
  const clientDetails = report?.client_details && typeof report.client_details === "object" ? report.client_details : {};
  const rawTitle = report?.report_title || buildExportTitle();
  const clientName =
    clientDetails.client_name ||
    clientDetails.name ||
    clientDetails.full_name ||
    clientDetails.preferred_name ||
    titleClientName(rawTitle) ||
    "Client";
  const reportDate =
    clientDetails.report_date ||
    clientDetails.visit_date ||
    clientDetails.assessment_date ||
    report?.report_date ||
    report?.date;
  const title = cleanReportTitle(rawTitle, clientName);
  return [
    safeFilenamePart(clientName, "Client"),
    safeFilenamePart(title, "Report"),
    safeFilenamePart(formatReportDate(reportDate, fallbackDate), "Date"),
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

function setReportPhotoStatus(message, isError = false) {
  if (!reportPhotoStatus) {
    return;
  }
  reportPhotoStatus.textContent = message;
  reportPhotoStatus.classList.toggle("error", isError);
}

function normalizePhotoClientName(value) {
  return String(value || "").trim();
}

function loadCachedReportPhotoClients() {
  try {
    const raw = localStorage.getItem(CLIENT_LIST_CACHE_KEY) || sessionStorage.getItem(CLIENT_LIST_CACHE_KEY);
    if (!raw) {
      return [];
    }
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed?.clients) || Date.now() - Number(parsed.cachedAt || 0) > CLIENT_LIST_CACHE_TTL_MS) {
      return [];
    }
    return parsed.clients;
  } catch {
    return [];
  }
}

function saveCachedReportPhotoClients(clients) {
  const payload = JSON.stringify({ clients, cachedAt: Date.now() });
  try {
    localStorage.setItem(CLIENT_LIST_CACHE_KEY, payload);
  } catch {
    // Ignore browser storage limits.
  }
  try {
    sessionStorage.setItem(CLIENT_LIST_CACHE_KEY, payload);
  } catch {
    // Ignore browser storage limits.
  }
}

function renderReportPhotoClientOptions() {
  if (!reportPhotoClientSelect) {
    return;
  }
  const current = reportPhotoClientSelect.value;
  reportPhotoClientSelect.innerHTML = "";
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = reportPhotoClients.length ? "Select client" : "No clients with images";
  reportPhotoClientSelect.append(placeholder);
  for (const client of reportPhotoClients.slice().sort((a, b) => a.localeCompare(b))) {
    const option = document.createElement("option");
    option.value = client;
    option.textContent = client;
    reportPhotoClientSelect.append(option);
  }
  reportPhotoClientSelect.value = reportPhotoClients.includes(current) ? current : "";
}

async function loadReportPhotoClients() {
  reportPhotoClients = loadCachedReportPhotoClients();
  renderReportPhotoClientOptions();
  try {
    const payload = await directoryApi.listMarketingPhotos({ clientsOnly: 1 });
    reportPhotoClients = (Array.isArray(payload?.clients) ? payload.clients : [])
      .map((item) => normalizePhotoClientName(item?.name))
      .filter((value, index, array) => Boolean(value) && array.indexOf(value) === index);
    saveCachedReportPhotoClients(reportPhotoClients);
    renderReportPhotoClientOptions();
  } catch (error) {
    setReportPhotoStatus(error?.message || "Could not load photo clients.", true);
  }
}

function reportPhotoClientFromReport() {
  if (structuredReport?.status === "ready_for_render") {
    const restoredReport = restoreReportPlaceholders(structuredReport);
    const details = restoredReport?.client_details || {};
    return normalizePhotoClientName(
      details.client_name || details.name || details.full_name || details.preferred_name || titleClientName(restoredReport.report_title)
    );
  }
  return normalizePhotoClientName(defaultPreferredName(buildEffectiveMapping()) || preferredNameInput.value);
}

async function loadReportPhotosForClient(clientName) {
  const client = normalizePhotoClientName(clientName);
  if (!client) {
    reportPhotoPool = [];
    renderReportPhotoImages();
    setReportPhotoStatus("Select a client or add local images.");
    return;
  }
  setReportPhotoStatus(`Loading images for ${client}...`);
  try {
    const payload = await directoryApi.listMarketingPhotos({ client });
    reportPhotoPool = (Array.isArray(payload?.photos) ? payload.photos : []).filter(isReportPhotoImage);
    renderReportPhotoImages();
    setReportPhotoStatus(`${reportPhotoPool.length} image${reportPhotoPool.length === 1 ? "" : "s"} found for ${client}.`);
  } catch (error) {
    setReportPhotoStatus(error?.message || "Could not load client photos.", true);
  }
}

function isReportPhotoImage(photo) {
  const type = String(photo?.mediaType || "").toLowerCase();
  const url = String(photo?.imageUrl || photo?.mediaUrl || photo?.attachmentUrl || "").toLowerCase();
  return type !== "video" && Boolean(url) && !/\.(mp4|mov|webm)(?:$|[?#])/.test(url);
}

function selectReportPhoto(photo) {
  const existing = reportPhotoSelected.findIndex((item) => item.id === photo.id);
  if (existing >= 0) {
    removeReportPhotoSelection(existing);
    return;
  }
  if (reportPhotoSelected.length >= MAX_SELECTION) {
    setReportPhotoStatus(`Select up to ${MAX_SELECTION} images.`, true);
    return;
  }
  reportPhotoSelected.push({
    id: photo.id,
    title: photo.title || photo.fileName || photo.client || "Untitled image",
    sourceCandidates: [photo.attachmentUrl, photo.mediaUrl, photo.imageUrl]
      .map((value) => String(value || "").trim())
      .filter((value, index, array) => Boolean(value) && array.indexOf(value) === index),
    previewUrl: photo.imageUrl || photo.mediaUrl || photo.attachmentUrl || "",
    zoom: 1,
    panX: 0,
    panY: 0,
  });
  if (reportPhotoSelectedSlot < 0) {
    reportPhotoSelectedSlot = 0;
  }
  invalidateReportCollage();
  renderReportPhotoComposer();
}

function removeReportPhotoSelection(index) {
  const [removed] = reportPhotoSelected.splice(index, 1);
  if (removed?.localObjectUrl) {
    URL.revokeObjectURL(removed.localObjectUrl);
  }
  if (reportPhotoSelectedSlot >= reportPhotoSelected.length) {
    reportPhotoSelectedSlot = reportPhotoSelected.length - 1;
  }
  invalidateReportCollage();
  renderReportPhotoComposer();
}

function addReportLocalImages(files = []) {
  const candidates = Array.from(files || []).filter((file) => String(file?.type || "").toLowerCase().startsWith("image/"));
  const accepted = candidates.slice(0, Math.max(0, MAX_SELECTION - reportPhotoSelected.length));
  for (const file of accepted) {
    const objectUrl = URL.createObjectURL(file);
    reportPhotoSelected.push({
      id: `report-local-${Date.now()}-${Math.random().toString(36).slice(2)}`,
      title: file.name || "Local image",
      sourceCandidates: [objectUrl],
      previewUrl: objectUrl,
      localObjectUrl: objectUrl,
      zoom: 1,
      panX: 0,
      panY: 0,
    });
  }
  if (reportPhotoSelectedSlot < 0 && reportPhotoSelected.length) {
    reportPhotoSelectedSlot = 0;
  }
  invalidateReportCollage();
  renderReportPhotoComposer();
  setReportPhotoStatus(accepted.length ? `Added ${accepted.length} local image${accepted.length === 1 ? "" : "s"}.` : "No images added.");
}

function renderReportPhotoImages() {
  if (!reportPhotoImagesGrid) {
    return;
  }
  reportPhotoImagesGrid.innerHTML = "";
  const selected = new Set(reportPhotoSelected.map((item) => item.id));
  for (const photo of reportPhotoPool) {
    const selectedPhoto = selected.has(photo.id);
    const button = document.createElement("button");
    button.type = "button";
    button.className = `layout-image-card${selectedPhoto ? " selected" : ""}`;
    button.disabled = !selectedPhoto && reportPhotoSelected.length >= MAX_SELECTION;
    button.setAttribute("aria-pressed", selectedPhoto ? "true" : "false");
    const img = document.createElement("img");
    img.className = "layout-image-thumb";
    img.src = photo.imageUrl || photo.mediaUrl || photo.attachmentUrl;
    img.alt = photo.title || photo.client || "Client image";
    img.loading = "lazy";
    const caption = document.createElement("span");
    caption.className = "layout-image-caption";
    caption.textContent = photo.title || photo.fileName || "Untitled image";
    button.append(img, caption);
    button.addEventListener("click", () => selectReportPhoto(photo));
    reportPhotoImagesGrid.append(button);
  }
}

function getReportPhotoLayout() {
  return LAYOUTS.find((layout) => layout.id === reportPhotoActiveLayoutId) || LAYOUTS[0];
}

function renderReportPhotoLayoutPicker() {
  if (!reportPhotoLayoutPicker) {
    return;
  }
  reportPhotoLayoutPicker.innerHTML = "";
  const selectedCount = reportPhotoSelected.length;
  const exact = selectedCount ? LAYOUTS.filter((layout) => layout.slots.length === selectedCount) : LAYOUTS.slice(0, 6);
  const visibleLayouts = exact.length ? exact : LAYOUTS;
  if (!visibleLayouts.some((layout) => layout.id === reportPhotoActiveLayoutId)) {
    reportPhotoActiveLayoutId = visibleLayouts[0].id;
  }
  for (const layout of visibleLayouts) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `layout-thumb${layout.id === reportPhotoActiveLayoutId ? " active" : ""}`;
    button.setAttribute("aria-pressed", layout.id === reportPhotoActiveLayoutId ? "true" : "false");
    const mini = document.createElement("div");
    mini.className = "layout-thumb-canvas";
    mini.style.setProperty("--layout-thumb-aspect", String(layout.aspect));
    for (const slot of layout.slots) {
      const cell = document.createElement("span");
      cell.className = "layout-thumb-slot";
      cell.style.left = `${slot.x * 100}%`;
      cell.style.top = `${slot.y * 100}%`;
      cell.style.width = `${slot.w * 100}%`;
      cell.style.height = `${slot.h * 100}%`;
      mini.append(cell);
    }
    const label = document.createElement("span");
    label.className = "layout-thumb-label";
    label.textContent = layout.name;
    button.append(mini, label);
    button.addEventListener("click", () => {
      reportPhotoActiveLayoutId = layout.id;
      reportPhotoSelectedSlot = Math.min(reportPhotoSelectedSlot, layout.slots.length - 1);
      invalidateReportCollage();
      renderReportPhotoComposer();
    });
    reportPhotoLayoutPicker.append(button);
  }
}

function getReportPhotoGapPx() {
  return reportPhotoStyle.gapEnabled ? clamp(reportPhotoStyle.gapPx, 0, 120) : 0;
}

function getReportPhotoCornerRadiusPx() {
  return reportPhotoStyle.roundedEnabled ? clamp(reportPhotoStyle.cornerRadiusPx, 0, 240) : 0;
}

function getReportPhotoBackgroundColor() {
  const color = String(reportPhotoStyle.backgroundColor || "").trim().toLowerCase();
  return LAYOUT_BACKGROUND_COLORS.has(color) ? color : "#ffffff";
}

function getReportPhotoSourceDimensions(previewUrl, fallbackWidth, fallbackHeight) {
  return reportPhotoImageMetaCache.get(String(previewUrl || "")) || {
    width: Math.max(1, Number(fallbackWidth) || 1),
    height: Math.max(1, Number(fallbackHeight) || 1),
  };
}

function renderReportPhotoStage() {
  if (!reportPhotoStage) {
    return;
  }
  const layout = getReportPhotoLayout();
  reportPhotoStage.innerHTML = "";
  reportPhotoStage.style.setProperty("--layout-stage-aspect", String(layout.aspect));
  const stageWidth = reportPhotoStage.clientWidth || reportPhotoStage.getBoundingClientRect().width || 1;
  const stageHeight = reportPhotoStage.clientHeight || stageWidth / layout.aspect || 1;
  const previewScale = stageWidth / EXPORT_WIDTH;
  const backgroundColor = getReportPhotoBackgroundColor();
  const nonWhiteBackground = backgroundColor !== "#ffffff";
  const gapPx = getReportPhotoGapPx() * previewScale;
  const cornerRadiusPx = getReportPhotoCornerRadiusPx() * previewScale;
  const outerPaddingPx = nonWhiteBackground ? gapPx : 0;
  const innerWidth = Math.max(1, stageWidth - outerPaddingPx * 2);
  const innerHeight = Math.max(1, stageHeight - outerPaddingPx * 2);
  reportPhotoStage.style.background = nonWhiteBackground ? backgroundColor : "#e9eff8";
  reportPhotoStage.style.borderRadius = `${nonWhiteBackground ? cornerRadiusPx : 8}px`;

  layout.slots.forEach((slot, slotIndex) => {
    const assignedImage = reportPhotoSelected[slotIndex];
    const neighborFlags = getSlotNeighborFlags(layout.slots, slotIndex);
    const base = computeAdjustedSlotRect(slot, neighborFlags, innerWidth, innerHeight, gapPx);
    const slotRect = { x: base.x + outerPaddingPx, y: base.y + outerPaddingPx, w: base.w, h: base.h };
    const slotEl = document.createElement("div");
    slotEl.className = `layout-slot${slotIndex === reportPhotoSelectedSlot ? " active" : ""}${assignedImage ? " has-image" : ""}`;
    slotEl.style.left = `${slotRect.x}px`;
    slotEl.style.top = `${slotRect.y}px`;
    slotEl.style.width = `${slotRect.w}px`;
    slotEl.style.height = `${slotRect.h}px`;
    slotEl.style.borderRadius = `${cornerRadiusPx}px`;
    slotEl.addEventListener("click", () => {
      reportPhotoSelectedSlot = slotIndex;
      renderReportPhotoComposer();
    });

    if (assignedImage) {
      const img = document.createElement("img");
      img.className = "layout-slot-image";
      img.src = assignedImage.previewUrl;
      img.alt = assignedImage.title;
      img.draggable = false;
      img.style.position = "absolute";
      const source = getReportPhotoSourceDimensions(assignedImage.previewUrl, slotRect.w, slotRect.h);
      const placement = computeImagePlacement(slotRect.w, slotRect.h, source.width, source.height, assignedImage.zoom, assignedImage.panX, assignedImage.panY);
      img.style.left = `${placement.renderX}px`;
      img.style.top = `${placement.renderY}px`;
      img.style.width = `${placement.renderW}px`;
      img.style.height = `${placement.renderH}px`;
      img.addEventListener("load", () => {
        if (img.naturalWidth > 0 && img.naturalHeight > 0) {
          const previous = reportPhotoImageMetaCache.get(assignedImage.previewUrl);
          if (!previous || previous.width !== img.naturalWidth || previous.height !== img.naturalHeight) {
            reportPhotoImageMetaCache.set(assignedImage.previewUrl, { width: img.naturalWidth, height: img.naturalHeight });
            renderReportPhotoStage();
          }
        }
      });
      img.addEventListener("pointerdown", (event) => {
        reportPhotoSelectedSlot = slotIndex;
        reportPhotoDragState = {
          slotIndex,
          pointerId: event.pointerId,
          startX: event.clientX,
          startY: event.clientY,
          startPanX: assignedImage.panX,
          startPanY: assignedImage.panY,
          maxOffsetX: placement.maxOffsetX,
          maxOffsetY: placement.maxOffsetY,
        };
        img.setPointerCapture(event.pointerId);
      });
      img.addEventListener("pointermove", (event) => {
        if (!reportPhotoDragState || reportPhotoDragState.pointerId !== event.pointerId || reportPhotoDragState.slotIndex !== slotIndex) {
          return;
        }
        const image = reportPhotoSelected[slotIndex];
        const deltaX = event.clientX - reportPhotoDragState.startX;
        const deltaY = event.clientY - reportPhotoDragState.startY;
        image.panX = clamp(reportPhotoDragState.startPanX + (reportPhotoDragState.maxOffsetX > 0 ? (deltaX / reportPhotoDragState.maxOffsetX) * 100 : 0), -100, 100);
        image.panY = clamp(reportPhotoDragState.startPanY + (reportPhotoDragState.maxOffsetY > 0 ? (deltaY / reportPhotoDragState.maxOffsetY) * 100 : 0), -100, 100);
        invalidateReportCollage();
        renderReportPhotoStage();
        renderReportPhotoAdjustControls();
      });
      img.addEventListener("pointerup", () => {
        reportPhotoDragState = null;
      });
      img.addEventListener("pointercancel", () => {
        reportPhotoDragState = null;
      });
      slotEl.append(img);
    } else {
      const placeholder = document.createElement("span");
      placeholder.className = "layout-slot-placeholder";
      placeholder.textContent = `Slot ${slotIndex + 1}`;
      slotEl.append(placeholder);
    }
    reportPhotoStage.append(slotEl);
  });
}

function renderReportPhotoSelectedList() {
  if (!reportPhotoSelectedList) {
    return;
  }
  reportPhotoSelectedList.innerHTML = "";
  if (!reportPhotoSelected.length) {
    reportPhotoSelectedList.innerHTML = '<p class="muted">No images selected yet.</p>';
    return;
  }
  reportPhotoSelected.forEach((image, index) => {
    const row = document.createElement("div");
    row.className = `selected-image-row${index === reportPhotoSelectedSlot ? " active" : ""}`;
    const info = document.createElement("button");
    info.type = "button";
    info.className = "selected-image-info";
    info.textContent = `${index + 1}. ${image.title}`;
    info.addEventListener("click", () => {
      reportPhotoSelectedSlot = index;
      renderReportPhotoComposer();
    });
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "secondary selected-image-remove";
    remove.textContent = "Remove";
    remove.addEventListener("click", () => removeReportPhotoSelection(index));
    row.append(info, remove);
    reportPhotoSelectedList.append(row);
  });
}

function renderReportPhotoAdjustControls() {
  const image = reportPhotoSelected[reportPhotoSelectedSlot];
  const disabled = !image;
  for (const input of [reportPhotoZoomRange, reportPhotoPanXRange, reportPhotoPanYRange, reportPhotoResetBtn]) {
    if (input) {
      input.disabled = disabled;
    }
  }
  if (!image) {
    return;
  }
  reportPhotoZoomRange.value = String(image.zoom);
  reportPhotoPanXRange.value = String(Math.round(image.panX));
  reportPhotoPanYRange.value = String(Math.round(image.panY));
}

function renderReportPhotoStyleControls() {
  if (reportPhotoGapEnabled) {
    reportPhotoGapEnabled.checked = reportPhotoStyle.gapEnabled;
  }
  if (reportPhotoRoundedEnabled) {
    reportPhotoRoundedEnabled.checked = reportPhotoStyle.roundedEnabled;
  }
  if (reportPhotoGapRange) {
    reportPhotoGapRange.value = String(reportPhotoStyle.gapPx);
    reportPhotoGapRange.disabled = !reportPhotoStyle.gapEnabled;
  }
  if (reportPhotoCornerRange) {
    reportPhotoCornerRange.value = String(reportPhotoStyle.cornerRadiusPx);
    reportPhotoCornerRange.disabled = !reportPhotoStyle.roundedEnabled;
  }
  if (reportPhotoGapValue) {
    reportPhotoGapValue.textContent = `${Math.round(reportPhotoStyle.gapPx)}px`;
  }
  if (reportPhotoCornerValue) {
    reportPhotoCornerValue.textContent = `${Math.round(reportPhotoStyle.cornerRadiusPx)}px`;
  }
  if (reportPhotoBackgroundPicker) {
    const active = getReportPhotoBackgroundColor();
    reportPhotoBackgroundPicker.querySelectorAll(".layout-color-swatch[data-color]").forEach((button) => {
      const selected = String(button.getAttribute("data-color") || "").toLowerCase() === active;
      button.classList.toggle("active", selected);
      button.setAttribute("aria-pressed", selected ? "true" : "false");
    });
  }
  if (generateCollageBtn) {
    generateCollageBtn.disabled = reportPhotoSelected.length === 0;
  }
  if (clearCollageBtn) {
    clearCollageBtn.disabled = !reportPhotoCollage && reportPhotoSelected.length === 0;
  }
}

function renderReportPhotoComposer() {
  renderReportPhotoImages();
  renderReportPhotoLayoutPicker();
  renderReportPhotoSelectedList();
  renderReportPhotoStage();
  renderReportPhotoAdjustControls();
  renderReportPhotoStyleControls();
}

function invalidateReportCollage() {
  reportPhotoCollage = null;
  if (reportPhotoCollageUrl) {
    URL.revokeObjectURL(reportPhotoCollageUrl);
    reportPhotoCollageUrl = "";
  }
  if (reportCollagePreview) {
    reportCollagePreview.hidden = true;
    reportCollagePreview.removeAttribute("src");
  }
  renderReportPhotoStyleControls();
}

function decodeBase64ToBlob(base64, mimeType) {
  const binary = atob(String(base64 || ""));
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return new Blob([bytes], { type: mimeType || "image/jpeg" });
}

function loadImageElement(url) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error(`Could not load image from ${url}`));
    img.src = url;
  });
}

async function loadReportExportImage(sourceCandidates = []) {
  const candidates = Array.isArray(sourceCandidates) ? sourceCandidates.filter(Boolean) : [];
  let lastError = null;
  for (const candidate of candidates) {
    if (reportPhotoExportImageCache.has(candidate)) {
      return reportPhotoExportImageCache.get(candidate);
    }
    try {
      if (/^(blob:|data:|\.\/assets\/|\/assets\/)/i.test(candidate)) {
        const image = await loadImageElement(candidate);
        reportPhotoExportImageCache.set(candidate, image);
        return image;
      }
      const payload = await directoryApi.getMarketingMedia({ url: candidate });
      const blob = decodeBase64ToBlob(payload?.dataBase64, payload?.mimeType);
      const objectUrl = URL.createObjectURL(blob);
      try {
        const image = await loadImageElement(objectUrl);
        reportPhotoExportImageCache.set(candidate, image);
        return image;
      } finally {
        URL.revokeObjectURL(objectUrl);
      }
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError || new Error("Could not load image.");
}

async function buildReportPhotoCanvas() {
  if (!reportPhotoSelected.length) {
    throw new Error("Select at least one image.");
  }
  const layout = getReportPhotoLayout();
  const canvas = document.createElement("canvas");
  canvas.width = EXPORT_WIDTH;
  canvas.height = Math.round(EXPORT_WIDTH / layout.aspect);
  const ctx = canvas.getContext("2d");
  if (!ctx) {
    throw new Error("Could not create collage canvas.");
  }
  const backgroundColor = getReportPhotoBackgroundColor();
  const nonWhiteBackground = backgroundColor !== "#ffffff";
  const gapPx = getReportPhotoGapPx();
  const cornerRadiusPx = getReportPhotoCornerRadiusPx();
  const outerPaddingPx = nonWhiteBackground ? gapPx : 0;
  const innerWidth = Math.max(1, canvas.width - outerPaddingPx * 2);
  const innerHeight = Math.max(1, canvas.height - outerPaddingPx * 2);
  ctx.fillStyle = backgroundColor;
  if (nonWhiteBackground) {
    ctx.save();
    ctx.beginPath();
    drawRoundedRectPath(ctx, 0, 0, canvas.width, canvas.height, cornerRadiusPx);
    ctx.clip();
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.restore();
  } else {
    ctx.fillRect(0, 0, canvas.width, canvas.height);
  }
  const renderableSlots = Math.min(layout.slots.length, reportPhotoSelected.length);
  for (let index = 0; index < renderableSlots; index += 1) {
    const imageState = reportPhotoSelected[index];
    const image = await loadReportExportImage(imageState.sourceCandidates);
    const slot = layout.slots[index];
    const neighborFlags = getSlotNeighborFlags(layout.slots, index);
    const base = computeAdjustedSlotRect(slot, neighborFlags, innerWidth, innerHeight, gapPx);
    drawImageIntoSlot(ctx, image, { x: base.x + outerPaddingPx, y: base.y + outerPaddingPx, w: base.w, h: base.h }, imageState, cornerRadiusPx);
  }
  return canvas;
}

function canvasToBlob(canvas, type = "image/png") {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) {
        reject(new Error("Could not encode collage image."));
        return;
      }
      resolve(blob);
    }, type);
  });
}

function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || "").split(",")[1] || "");
    reader.onerror = () => reject(reader.error || new Error("Could not read image."));
    reader.readAsDataURL(blob);
  });
}

async function generateReportCollage() {
  try {
    setReportPhotoStatus("Generating collage...");
    const canvas = await buildReportPhotoCanvas();
    const blob = await canvasToBlob(canvas);
    reportPhotoCollage = {
      dataBase64: await blobToBase64(blob),
      mimeType: "image/png",
      width: canvas.width,
      height: canvas.height,
      layoutId: getReportPhotoLayout().id,
      imageCount: reportPhotoSelected.length,
    };
    if (reportPhotoCollageUrl) {
      URL.revokeObjectURL(reportPhotoCollageUrl);
    }
    reportPhotoCollageUrl = URL.createObjectURL(blob);
    if (reportCollagePreview) {
      reportCollagePreview.src = reportPhotoCollageUrl;
      reportCollagePreview.hidden = false;
    }
    setReportPhotoStatus("Collage ready for report export.");
    renderReportPhotoStyleControls();
  } catch (error) {
    setReportPhotoStatus(error?.message || "Could not generate collage.", true);
  }
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

function lineWidthForBlock(block) {
  if (block.type === "title") {
    return 54;
  }
  if (block.type === "heading") {
    return 68;
  }
  if (block.type === "bullet") {
    return 82;
  }
  return 88;
}

function pdfBlockStyle(block) {
  if (block.type === "title") {
    return { font: "F2", size: 18, lineHeight: 22, before: 0, after: 16, indent: 0 };
  }
  if (block.type === "heading") {
    return { font: "F2", size: 13, lineHeight: 17, before: 10, after: 5, indent: 0 };
  }
  if (block.type === "meta") {
    return { font: "F1", size: 10, lineHeight: 14, before: 0, after: 2, indent: 0 };
  }
  if (block.type === "bullet") {
    return { font: "F1", size: 10.5, lineHeight: 14, before: 0, after: 4, indent: 14 };
  }
  return { font: "F1", size: 10.5, lineHeight: 14, before: 0, after: 8, indent: 0 };
}

function textToPdfBlocks(title, text) {
  const blocks = [{ type: "title", text: title }];
  for (const rawBlock of normalizePdfText(text).split(/\n{2,}/)) {
    const trimmed = rawBlock.trim();
    if (!trimmed) {
      continue;
    }
    const [firstLine, ...rest] = trimmed.split(/\n/);
    if (rest.length > 0 && !firstLine.startsWith("- ")) {
      blocks.push({ type: "heading", text: firstLine });
      for (const line of rest) {
        blocks.push({ type: line.trim().startsWith("- ") ? "bullet" : "body", text: line.trim() });
      }
    } else {
      blocks.push({ type: trimmed.startsWith("- ") ? "bullet" : "body", text: trimmed });
    }
  }
  return blocks;
}

function buildPdfBlob({ title, text, blocks = null }) {
  const pageWidth = 595;
  const pageHeight = 842;
  const margin = 54;
  const pageEntries = [[]];
  let y = pageHeight - margin;

  const addPage = () => {
    pageEntries.push([]);
    y = pageHeight - margin;
  };

  for (const block of blocks || textToPdfBlocks(title, text)) {
    const value = normalizePdfText(block.text).trim();
    if (!value) {
      continue;
    }
    const style = pdfBlockStyle(block);
    y -= style.before;
    const lines = [];
    for (const line of value.split(/\n/)) {
      lines.push(...wrapPdfLine(line, lineWidthForBlock(block)));
    }
    for (const line of lines) {
      if (y - style.lineHeight < margin) {
        addPage();
      }
      pageEntries[pageEntries.length - 1].push({
        text: block.type === "bullet" && !line.startsWith("- ") ? `- ${line}` : line,
        x: margin + style.indent,
        y,
        font: style.font,
        size: style.size,
      });
      y -= style.lineHeight;
    }
    y -= style.after;
  }

  const objects = [];
  const addObject = (content) => {
    objects.push(content);
    return objects.length;
  };

  const catalogId = addObject("<< /Type /Catalog /Pages 2 0 R >>");
  const pagesId = addObject("");
  const fontId = addObject("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>");
  const boldFontId = addObject("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>");
  const pageIds = [];

  pageEntries.forEach((entries) => {
    const commands = entries.map(
      (entry) => `BT\n/${entry.font} ${entry.size} Tf\n${entry.x} ${entry.y} Td\n(${escapePdfText(entry.text)}) Tj\nET`
    );

    const stream = commands.join("\n");
    const contentId = addObject(`<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`);
    const pageId = addObject(
      `<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 ${pageWidth} ${pageHeight}] /Resources << /Font << /F1 ${fontId} 0 R /F2 ${boldFontId} 0 R >> >> /Contents ${contentId} 0 R >>`
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

function pushPdfSection(blocks, heading, value) {
  const text = stringifyReportValue(value);
  if (!text) {
    return;
  }
  blocks.push({ type: "heading", text: heading });
  blocks.push({ type: "body", text });
}

function pushPdfListSection(blocks, heading, items, suffix = "") {
  const values = (Array.isArray(items) ? items : []).map(stringifyReportValue).filter(Boolean);
  if (!values.length) {
    return;
  }
  blocks.push({ type: "heading", text: heading });
  values.forEach((item) => blocks.push({ type: "bullet", text: item }));
  if (suffix) {
    blocks.push({ type: "body", text: suffix });
  }
}

function buildStructuredReportPdfBlocks(report) {
  const sections = report.report_sections || {};
  const omitted = new Set((report.omitted_sections || []).map((item) => String(item || "").toLowerCase()));
  const blocks = [{ type: "title", text: report.report_title || buildExportTitle() }];
  for (const line of buildClientDetailsText(report.client_details).split("\n").filter(Boolean)) {
    blocks.push({ type: "meta", text: line });
  }
  pushPdfSection(blocks, "Executive Summary", sections.executive_summary);
  pushPdfSection(blocks, "1. Current Situation", sections.current_situation);
  pushPdfSection(blocks, "2. Physical Wellbeing", sections.physical_wellbeing);
  pushPdfSection(blocks, "3. Emotional Wellbeing", sections.emotional_wellbeing);
  pushPdfSection(blocks, "4. Environmental Wellbeing", sections.environmental_wellbeing);
  if (sections.wellbeing_highlights && !omitted.has("wellbeing highlights")) {
    pushPdfSection(blocks, "5. Wellbeing Highlights", sections.wellbeing_highlights);
  }
  pushPdfListSection(
    blocks,
    "Suggested SMART Goals",
    report.suggested_smart_goals,
    "These goals are suggested and can be reviewed or amended."
  );
  pushPdfListSection(blocks, "Recommendations to Improve Quality of Life", sections.recommendations);
  pushPdfListSection(blocks, "Next Steps", sections.next_steps);
  return blocks;
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
    downloadBlob(
      buildPdfBlob({
        title: restoredReport.report_title || buildExportTitle(),
        blocks: buildStructuredReportPdfBlocks(restoredReport),
      }),
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
      images: reportPhotoCollage ? { collage: reportPhotoCollage } : undefined,
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
    renderReportPhotoComposer();
    void loadReportPhotoClients();
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
showPlaceholderText?.addEventListener("change", syncDraftTextVisibility);
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
reportPhotoClientSelect?.addEventListener("change", () => {
  void loadReportPhotosForClient(reportPhotoClientSelect.value);
});
reportPhotoLoadBtn?.addEventListener("click", () => {
  const client = reportPhotoClientFromReport();
  if (client && reportPhotoClientSelect) {
    if (!reportPhotoClients.includes(client)) {
      reportPhotoClients = [...reportPhotoClients, client];
      renderReportPhotoClientOptions();
    }
    reportPhotoClientSelect.value = client;
  }
  void loadReportPhotosForClient(client);
});
reportPhotoLocalDropZone?.addEventListener("click", () => reportPhotoLocalInput?.click());
reportPhotoLocalDropZone?.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    reportPhotoLocalInput?.click();
  }
});
reportPhotoLocalDropZone?.addEventListener("dragover", (event) => {
  event.preventDefault();
  reportPhotoLocalDropZone.classList.add("is-dragover");
});
reportPhotoLocalDropZone?.addEventListener("dragleave", () => {
  reportPhotoLocalDropZone.classList.remove("is-dragover");
});
reportPhotoLocalDropZone?.addEventListener("drop", (event) => {
  event.preventDefault();
  reportPhotoLocalDropZone.classList.remove("is-dragover");
  addReportLocalImages(event.dataTransfer?.files || []);
});
reportPhotoLocalInput?.addEventListener("change", () => {
  addReportLocalImages(reportPhotoLocalInput.files || []);
  reportPhotoLocalInput.value = "";
});
reportPhotoZoomRange?.addEventListener("input", () => {
  const image = reportPhotoSelected[reportPhotoSelectedSlot];
  if (!image) {
    return;
  }
  image.zoom = clamp(Number(reportPhotoZoomRange.value) || 1, 1, 3);
  invalidateReportCollage();
  renderReportPhotoStage();
});
reportPhotoPanXRange?.addEventListener("input", () => {
  const image = reportPhotoSelected[reportPhotoSelectedSlot];
  if (!image) {
    return;
  }
  image.panX = clamp(Number(reportPhotoPanXRange.value) || 0, -100, 100);
  invalidateReportCollage();
  renderReportPhotoStage();
});
reportPhotoPanYRange?.addEventListener("input", () => {
  const image = reportPhotoSelected[reportPhotoSelectedSlot];
  if (!image) {
    return;
  }
  image.panY = clamp(Number(reportPhotoPanYRange.value) || 0, -100, 100);
  invalidateReportCollage();
  renderReportPhotoStage();
});
reportPhotoResetBtn?.addEventListener("click", () => {
  const image = reportPhotoSelected[reportPhotoSelectedSlot];
  if (!image) {
    return;
  }
  image.zoom = 1;
  image.panX = 0;
  image.panY = 0;
  invalidateReportCollage();
  renderReportPhotoComposer();
});
reportPhotoGapEnabled?.addEventListener("change", () => {
  reportPhotoStyle.gapEnabled = Boolean(reportPhotoGapEnabled.checked);
  invalidateReportCollage();
  renderReportPhotoComposer();
});
reportPhotoGapRange?.addEventListener("input", () => {
  reportPhotoStyle.gapPx = clamp(Number(reportPhotoGapRange.value) || 0, 0, 120);
  invalidateReportCollage();
  renderReportPhotoComposer();
});
reportPhotoRoundedEnabled?.addEventListener("change", () => {
  reportPhotoStyle.roundedEnabled = Boolean(reportPhotoRoundedEnabled.checked);
  invalidateReportCollage();
  renderReportPhotoComposer();
});
reportPhotoCornerRange?.addEventListener("input", () => {
  reportPhotoStyle.cornerRadiusPx = clamp(Number(reportPhotoCornerRange.value) || 0, 0, 240);
  invalidateReportCollage();
  renderReportPhotoComposer();
});
reportPhotoBackgroundPicker?.addEventListener("click", (event) => {
  const target = event.target instanceof Element ? event.target.closest(".layout-color-swatch[data-color]") : null;
  if (!target) {
    return;
  }
  const color = String(target.getAttribute("data-color") || "").toLowerCase();
  reportPhotoStyle.backgroundColor = LAYOUT_BACKGROUND_COLORS.has(color) ? color : "#ffffff";
  invalidateReportCollage();
  renderReportPhotoComposer();
});
generateCollageBtn?.addEventListener("click", () => {
  void generateReportCollage();
});
clearCollageBtn?.addEventListener("click", () => {
  reportPhotoSelected.forEach((image) => {
    if (image?.localObjectUrl) {
      URL.revokeObjectURL(image.localObjectUrl);
    }
  });
  reportPhotoSelected = [];
  reportPhotoPool = [];
  reportPhotoSelectedSlot = -1;
  invalidateReportCollage();
  renderReportPhotoComposer();
  setReportPhotoStatus("Collage cleared.");
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

window.addEventListener("beforeunload", () => {
  for (const image of reportPhotoSelected) {
    if (image?.localObjectUrl) {
      URL.revokeObjectURL(image.localObjectUrl);
    }
  }
  if (reportPhotoCollageUrl) {
    URL.revokeObjectURL(reportPhotoCollageUrl);
  }
});

updateCounts();
renderReviewOutput();
applyReportModeDefaults();
void init();
