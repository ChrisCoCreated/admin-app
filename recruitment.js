import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260409";

const searchInput = document.getElementById("searchInput");
const locationFilterSelect = document.getElementById("locationFilterSelect");
const statusFilterSelect = document.getElementById("statusFilterSelect");
const sourceFilterSelect = document.getElementById("sourceFilterSelect");
const ownerFilterSelect = document.getElementById("ownerFilterSelect");
const activeFilterSelect = document.getElementById("activeFilterSelect");
const sortFilterSelect = document.getElementById("sortFilterSelect");
const mineOnlyFilterInput = document.getElementById("mineOnlyFilterInput");
const stageModeFilterButtons = Array.from(document.querySelectorAll("[data-stage-mode-filter]"));
const sortHeaderButtons = Array.from(document.querySelectorAll("[data-sort-header]"));
const recruitmentTableBody = document.getElementById("recruitmentTableBody");
const emptyState = document.getElementById("emptyState");
const statusMessage = document.getElementById("statusMessage");
const signOutBtn = document.getElementById("signOutBtn");
const detailRoot = document.getElementById("candidateDetail");
const sharePointListLink = document.getElementById("sharePointListLink");
const openIndeedBtn = document.getElementById("openIndeedBtn");
const openOneTouchBtn = document.getElementById("openOneTouchBtn");
const statusUpdateSelect = document.getElementById("statusUpdateSelect");
const saveStatusBtn = document.getElementById("saveStatusBtn");
const saveDetailBtn = document.getElementById("saveDetailBtn");
const importDropZone = document.getElementById("importDropZone");
const importFileInput = document.getElementById("importFileInput");
const importFileName = document.getElementById("importFileName");
const importSummary = document.getElementById("importSummary");
const importErrors = document.getElementById("importErrors");
const runImportBtn = document.getElementById("runImportBtn");
const importPreviewWrap = document.getElementById("importPreviewWrap");
const importPreviewTitle = document.getElementById("importPreviewTitle");
const importPreviewBody = document.getElementById("importPreviewBody");
const toggleRecruitmentToolbarBtn = document.getElementById("toggleRecruitmentToolbarBtn");
const addRecruitmentItemBtn = document.getElementById("addRecruitmentItemBtn");
const recruitmentToolbarContent = document.getElementById("recruitmentToolbarContent");
const candidateDetailModal = document.getElementById("candidateDetailModal");
const candidateDetailModalTitle = document.getElementById("candidateDetailModalTitle");
const candidateDetailCloseBtn = document.getElementById("candidateDetailCloseBtn");
const candidateDetailEditForm = document.getElementById("candidateDetailEditForm");
const addRecruitmentModal = document.getElementById("addRecruitmentModal");
const addRecruitmentCloseBtn = document.getElementById("addRecruitmentCloseBtn");
const addRecruitmentForm = document.getElementById("addRecruitmentForm");
const addRecruitmentError = document.getElementById("addRecruitmentError");
const addRecruitmentJsonDropZone = document.getElementById("addRecruitmentJsonDropZone");
const addRecruitmentJsonFileInput = document.getElementById("addRecruitmentJsonFileInput");
const addRecruitmentJsonFileName = document.getElementById("addRecruitmentJsonFileName");
const addCandidateNameInput = document.getElementById("addCandidateNameInput");
const addCandidateStatusSelect = document.getElementById("addCandidateStatusSelect");
const addCandidateEmailInput = document.getElementById("addCandidateEmailInput");
const addCandidatePhoneInput = document.getElementById("addCandidatePhoneInput");
const addCandidateLivesInInput = document.getElementById("addCandidateLivesInInput");
const addCandidateJobLocationInput = document.getElementById("addCandidateJobLocationInput");
const addCandidateSourceInput = document.getElementById("addCandidateSourceInput");
const addCandidateIndeedUrlInput = document.getElementById("addCandidateIndeedUrlInput");
const addCandidateActiveInput = document.getElementById("addCandidateActiveInput");
const addCandidateUpdateExistingInput = document.getElementById("addCandidateUpdateExistingInput");
const addCandidateNotesInput = document.getElementById("addCandidateNotesInput");
const saveRecruitmentCandidateBtn = document.getElementById("saveRecruitmentCandidateBtn");
const cancelAddRecruitmentBtn = document.getElementById("cancelAddRecruitmentBtn");
const oneTouchPickerModal = document.getElementById("oneTouchPickerModal");
const oneTouchPickerCandidate = document.getElementById("oneTouchPickerCandidate");
const oneTouchAreaSelect = document.getElementById("oneTouchAreaSelect");
const oneTouchRecruitmentSourceSelect = document.getElementById("oneTouchRecruitmentSourceSelect");
const oneTouchPositionSelect = document.getElementById("oneTouchPositionSelect");
const oneTouchStatusSelect = document.getElementById("oneTouchStatusSelect");
const oneTouchPickerError = document.getElementById("oneTouchPickerError");
const oneTouchPickerConfirmBtn = document.getElementById("oneTouchPickerConfirmBtn");
const oneTouchPickerCancelBtn = document.getElementById("oneTouchPickerCancelBtn");
const statusQuickMenu = document.getElementById("statusQuickMenu");
const statusQuickMenuList = document.getElementById("statusQuickMenuList");

const detailFields = {
  candidateName: detailRoot?.querySelector('[data-field="candidateName"]'),
  location: detailRoot?.querySelector('[data-field="location"]'),
  status: detailRoot?.querySelector('[data-field="status"]'),
  active: detailRoot?.querySelector('[data-field="active"]'),
  source: detailRoot?.querySelector('[data-field="source"]'),
  phoneNumber: detailRoot?.querySelector('[data-field="phoneNumber"]'),
  email: detailRoot?.querySelector('[data-field="email"]'),
  tags: detailRoot?.querySelector('[data-field="tags"]'),
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

const detailInputs = {
  candidateName: document.getElementById("detailCandidateNameInput"),
  location: document.getElementById("detailLocationInput"),
  source: document.getElementById("detailSourceInput"),
  phoneNumber: document.getElementById("detailPhoneInput"),
  email: document.getElementById("detailEmailInput"),
  indeedUrl: document.getElementById("detailIndeedUrlInput"),
  livesIn: document.getElementById("detailLivesInInput"),
  earmarkedFor: document.getElementById("detailEarmarkedForInput"),
  keepInMind: document.getElementById("detailKeepInMindInput"),
  tags: document.getElementById("detailTagsInput"),
  notes: document.getElementById("detailNotesInput"),
};

const detailTagsPreview = document.getElementById("detailTagsPreview");

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
let statusUpdateBusy = false;
let activeUpdateBusy = false;
let ownerUpdateBusy = false;
let statusQuickMenuCandidateId = "";
let createCandidateBusy = false;
let activeHideRefreshTimer = 0;
let detailSaveBusy = false;
let stageUpdateBusyKey = "";
let statusFeedbackTimer = 0;
let stageModeFilter = "all";
let stageModeCountRenderToken = 0;
let recruitmentStatusOptions = [];
let recruitmentOwnerOptions = [];
let currentUserEmail = "";
let currentUserOwnerChoice = "";
const pendingInactiveReviewIds = new Set();
const dismissedInactiveReviewIds = new Set();
const openStageKeys = new Set();
const ONE_TOUCH_DEFAULT_AREA = "East Kent";
const ONE_TOUCH_DEFAULT_POSITION = "Health & Wellbeing Associate";
const ONE_TOUCH_DEFAULT_STATUS = "Pending";
const STATUS_FILTER_DEFAULT = "__default__";
const ONE_TOUCH_ELIGIBLE_STATUSES = new Set([
  "2nd Interview",
  "Exploring an offer",
  "Make Offer",
  "Offered",
  "Accepted",
  "Start Date Agreed",
  "Started",
]);
const DEFAULT_RECRUITMENT_STATUS_OPTIONS = [
  "Organise Initial Call",
  "Initial Call",
  "Organise 1st Interview",
  "1st Interview",
  "Organise 2nd Interview",
  "2nd Interview",
  "Exploring an offer",
  "Make Offer",
  "Offered",
  "Accepted",
  "Start Date Agreed",
  "Started",
  "Lost",
];
const DEFAULT_RECRUITMENT_OWNER_OPTIONS = [
  { label: "Chris" },
  { label: "Rebecca" },
  { label: "Miska" },
  { label: "Peter" },
];
recruitmentStatusOptions = [...DEFAULT_RECRUITMENT_STATUS_OPTIONS];
recruitmentOwnerOptions = [...DEFAULT_RECRUITMENT_OWNER_OPTIONS];
const INITIAL_STAGE_STATUSES = new Set(["Organise Initial Call", "Initial Call"]);
const INTERVIEWING_STAGE_STATUSES = new Set([
  "Organise 1st Interview",
  "1st Interview",
  "Organise 2nd Interview",
  "2nd Interview",
]);
const ONBOARDING_STAGE_STATUSES = new Set([
  "Exploring an offer",
  "Make Offer",
  "Offered",
  "Accepted",
  "Start Date Agreed",
  "Started",
]);

function normalizeText(value) {
  return String(value || "").trim().toLowerCase();
}

function cleanText(value) {
  return String(value || "").trim();
}

function getFirstNameSortValue(candidateName) {
  const fullName = cleanText(candidateName);
  if (!fullName) {
    return "";
  }
  return cleanText(fullName.split(/\s+/).find(Boolean) || "").toLowerCase();
}

function normalizeStatusOptions(options) {
  const list = Array.isArray(options) ? options.map(cleanText).filter(Boolean) : [];
  return list.length ? Array.from(new Set(list)) : [...DEFAULT_RECRUITMENT_STATUS_OPTIONS];
}

function normalizeOwnerOptions(options) {
  const list = Array.isArray(options) ? options.map(cleanText).filter(Boolean) : [];
  const labels = list.length ? Array.from(new Set(list)) : DEFAULT_RECRUITMENT_OWNER_OPTIONS.map((option) => option.label);
  return labels.map((label) => ({ label }));
}

function syncSortHeaderButtons() {
  const selectedSort = cleanText(sortFilterSelect?.value || "updated_desc");
  for (const button of sortHeaderButtons) {
    const headerKey = cleanText(button?.dataset?.sortHeader);
    let direction = "";
    if (headerKey === "candidate") {
      direction = selectedSort === "name_asc" ? "asc" : selectedSort === "name_desc" ? "desc" : "";
    } else if (headerKey === "location") {
      direction = selectedSort === "location_asc" ? "asc" : selectedSort === "location_desc" ? "desc" : "";
    } else if (headerKey === "status") {
      direction = selectedSort === "status_asc" ? "asc" : selectedSort === "status_desc" ? "desc" : "";
    }
    button.dataset.direction = direction;
    button.classList.toggle("is-active", Boolean(direction));
  }
}

function toggleHeaderSort(headerKey) {
  if (!sortFilterSelect) {
    return;
  }
  const selectedSort = cleanText(sortFilterSelect.value || "updated_desc");
  let nextSort = "updated_desc";
  if (headerKey === "candidate") {
    nextSort = selectedSort === "name_asc" ? "name_desc" : "name_asc";
  } else if (headerKey === "location") {
    nextSort = selectedSort === "location_asc" ? "location_desc" : "location_asc";
  } else if (headerKey === "status") {
    nextSort = selectedSort === "status_asc" ? "status_desc" : "status_asc";
  }
  sortFilterSelect.value = nextSort;
  syncSortHeaderButtons();
  renderCandidates();
}

function normalizeComparable(value) {
  return cleanText(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function resolveMineOwnerChoice(email, options) {
  const normalizedEmail = cleanText(email).toLowerCase();
  if (!normalizedEmail) {
    return "";
  }
  const explicitCandidates = {
    "chris@planwithcare.co.uk": ["chris"],
    "rebecca@planwithcare.co.uk": ["rebecca"],
    "peter@planwithcare.co.uk": ["peter"],
    "michalina@thrivehomecare.co.uk": ["miska", "miśka", "michalina"],
  };
  const candidates = explicitCandidates[normalizedEmail] || [normalizedEmail.split("@")[0] || ""];
  const option = options.find((entry) => {
    const comparable = normalizeComparable(entry?.label);
    return candidates.some((candidate) => {
      const target = normalizeComparable(candidate);
      return target && (comparable === target || comparable.includes(target) || target.includes(comparable));
    });
  });
  return cleanText(option?.label);
}

function getCurrentOwnerLabel(candidate) {
  const owner = cleanText(candidate?.currentOwner);
  const matched = recruitmentOwnerOptions.find((option) => cleanText(option.label).toLowerCase() === owner.toLowerCase());
  return matched?.label || owner || "Unassigned";
}

function setStageModeFilter(nextValue = "all") {
  const normalized = cleanText(nextValue);
  stageModeFilter = ["all", "initial", "interviewing", "onboarding", "hold", "reject"].includes(normalized)
    ? normalized
    : "all";
  for (const button of stageModeFilterButtons) {
    const isActive = cleanText(button?.dataset?.stageModeFilter) === stageModeFilter;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-pressed", isActive ? "true" : "false");
  }
}

function getStageModeCount(mode, candidates) {
  if (mode === "all") {
    return Array.isArray(candidates) ? candidates.length : 0;
  }
  const previousMode = stageModeFilter;
  stageModeFilter = mode;
  const count = (Array.isArray(candidates) ? candidates : []).filter((candidate) => candidateMatchesStageMode(candidate)).length;
  stageModeFilter = previousMode;
  return count;
}

function updateStageModeFilterButtons(candidates) {
  for (const button of stageModeFilterButtons) {
    const mode = cleanText(button?.dataset?.stageModeFilter);
    const baseLabel = cleanText(button?.dataset?.label || button?.textContent);
    const count = getStageModeCount(mode, candidates);
    button.dataset.label = baseLabel;
    button.dataset.count = String(count);
    button.textContent = `${baseLabel} ${count}`;
    button.classList.toggle("has-items", count > 0);
    button.classList.toggle("has-items-hold", mode === "hold" && count > 0);
    button.classList.toggle("has-items-reject", mode === "reject" && count > 0);
  }
}

function getCandidateProcessOutcomes(candidate) {
  return [
    cleanText(candidate?.screenOutcome),
    cleanText(candidate?.firstInterviewOutcome),
    cleanText(candidate?.secondInterviewOutcome),
  ].filter(Boolean);
}

function candidateMatchesStageMode(candidate) {
  const status = cleanText(candidate?.status);
  if (stageModeFilter === "initial") {
    return INITIAL_STAGE_STATUSES.has(status);
  }
  if (stageModeFilter === "interviewing") {
    return INTERVIEWING_STAGE_STATUSES.has(status);
  }
  if (stageModeFilter === "onboarding") {
    return ONBOARDING_STAGE_STATUSES.has(status);
  }
  if (stageModeFilter === "hold") {
    return getCandidateProcessOutcomes(candidate).some((outcome) => normalizeText(outcome) === "hold");
  }
  if (stageModeFilter === "reject") {
    return getCandidateProcessOutcomes(candidate).some((outcome) => normalizeText(outcome) === "reject");
  }
  return true;
}

function getBaseFilteredCandidates() {
  const query = normalizeText(searchInput.value);
  const selectedLocation = cleanText(locationFilterSelect.value || "all");
  const selectedStatus = cleanText(statusFilterSelect.value || STATUS_FILTER_DEFAULT);
  const selectedSource = cleanText(sourceFilterSelect.value || "all");
  const selectedOwner = cleanText(ownerFilterSelect?.value || "all");
  const selectedActive = cleanText(activeFilterSelect?.value || "active");
  const mineOnly = mineOnlyFilterInput?.checked === true;

  return allCandidates.filter((candidate) => {
    const candidateId = cleanText(candidate?.id);
    if (dismissedInactiveReviewIds.has(candidateId)) {
      return false;
    }
    if (selectedActive === "active" && candidate.active !== true && !pendingInactiveReviewIds.has(candidateId)) {
      return false;
    }
    if (selectedActive === "inactive" && candidate.active !== false) {
      return false;
    }
    if (selectedLocation !== "all" && cleanText(candidate.location) !== selectedLocation) {
      return false;
    }
    if (selectedStatus === STATUS_FILTER_DEFAULT && normalizeText(candidate.status) === "rejected") {
      return false;
    }
    if (
      selectedStatus !== "all" &&
      selectedStatus !== STATUS_FILTER_DEFAULT &&
      cleanText(candidate.status) !== selectedStatus
    ) {
      return false;
    }
    if (selectedSource !== "all" && cleanText(candidate.source) !== selectedSource) {
      return false;
    }
    if (selectedOwner !== "all" && cleanText(candidate.currentOwner) !== selectedOwner) {
      return false;
    }
    if (mineOnly && (!currentUserOwnerChoice || cleanText(candidate.currentOwner) !== currentUserOwnerChoice)) {
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
      normalizeText(candidate.notes).includes(query) ||
      normalizeText(candidate.tags).includes(query)
    );
  });
}

function getStageTone(value) {
  const normalized = normalizeText(value);
  if (!normalized) {
    return "neutral";
  }
  if (
    normalized.includes("reject") ||
    normalized.includes("red") ||
    normalized.includes("fail") ||
    normalized.includes("no")
  ) {
    return "negative";
  }
  if (
    normalized.includes("progress") ||
    normalized.includes("accept") ||
    normalized.includes("offer") ||
    normalized.includes("start") ||
    normalized.includes("pass") ||
    normalized.includes("yes")
  ) {
    return "positive";
  }
  if (
    normalized.includes("hold") ||
    normalized.includes("pending") ||
    normalized.includes("review") ||
    normalized.includes("arrange") ||
    normalized.includes("organise")
  ) {
    return "pending";
  }
  return "neutral";
}

function renderStageSummary(candidate) {
  const stages = [
    {
      key: "initial_screen",
      label: "Initial Screen",
      outcome: cleanText(candidate?.screenOutcome),
      nextSteps: cleanText(candidate?.screenNextSteps),
    },
    {
      key: "first_interview",
      label: "1st Interview",
      outcome: cleanText(candidate?.firstInterviewOutcome),
      nextSteps: cleanText(candidate?.firstInterviewNextSteps),
    },
    {
      key: "second_interview",
      label: "2nd Interview",
      outcome: cleanText(candidate?.secondInterviewOutcome),
      nextSteps: cleanText(candidate?.secondInterviewNextSteps),
    },
  ];

  return `
    <div class="recruitment-stage-grid">
      ${stages
        .map((stage) => {
          const outcome = stage.outcome || "Not recorded";
          const tone = getStageTone(stage.outcome);
          const isEmpty = !stage.outcome && !stage.nextSteps;
          const nextSteps = stage.nextSteps || "";
          const busyKey = `${cleanText(candidate?.id)}:${stage.key}`;
          const isBusy = stageUpdateBusyKey === busyKey;
          const openKey = `${cleanText(candidate?.id)}:${stage.key}`;
          return `
            <details class="recruitment-stage-card recruitment-stage-card-${tone}${isEmpty ? " recruitment-stage-card-empty" : ""}" data-item-id="${escapeHtml(
              cleanText(candidate?.id)
            )}" data-stage-key="${escapeHtml(stage.key)}"${openStageKeys.has(openKey) ? " open" : ""}>
              <summary class="recruitment-stage-summary">
                <span class="recruitment-stage-summary-main">
                  <span class="recruitment-stage-label">${escapeHtml(stage.label)}</span>
                  <span class="recruitment-stage-outcome">${escapeHtml(outcome)}</span>
                </span>
                <span class="recruitment-stage-summary-meta">${isEmpty ? "Add" : nextSteps ? "Notes" : "View"}</span>
              </summary>
              <div class="recruitment-stage-editor" data-stage-editor>
                <label class="field compact-field">
                  Outcome
                  <select data-stage-outcome>
                    <option value="">Not recorded</option>
                    <option value="Progress"${stage.outcome === "Progress" ? " selected" : ""}>Progress</option>
                    <option value="Hold"${stage.outcome === "Hold" ? " selected" : ""}>Hold</option>
                    <option value="Reject"${stage.outcome === "Reject" ? " selected" : ""}>Reject</option>
                  </select>
                </label>
                <label class="field compact-field">
                  Next Steps
                  <textarea data-stage-next-steps placeholder="Add notes or next steps...">${escapeHtml(nextSteps)}</textarea>
                </label>
                <div class="recruitment-stage-editor-actions">
                  <button type="button" class="secondary recruitment-stage-save-btn" data-stage-save${isBusy ? " disabled" : ""}>${
                    isBusy ? "Saving..." : "Save"
                  }</button>
                </div>
              </div>
            </details>
          `;
        })
        .join("")}
    </div>
  `;
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function splitTagList(value) {
  return Array.from(
    new Set(
      String(value || "")
        .split(/[,\n;|]+/)
        .map((part) => cleanText(part))
        .filter(Boolean)
    )
  );
}

function normalizeTagString(value) {
  return splitTagList(value).join(", ");
}

function renderTagPreview(node, tags, emptyLabel = "No tags yet") {
  if (!node) {
    return;
  }
  const list = Array.isArray(tags) ? tags : splitTagList(tags);
  if (!list.length) {
    node.innerHTML = `<span class="tag-chip tag-chip-muted">${escapeHtml(emptyLabel)}</span>`;
    return;
  }
  node.innerHTML = list.map((tag) => `<span class="tag-chip">${escapeHtml(tag)}</span>`).join("");
}

function setStatus(message, isError = false, options = {}) {
  if (!statusMessage) {
    return;
  }
  const text = cleanText(message);
  const subtle = options.subtle === true && !isError;
  const autoClear = options.autoClear !== false && !isError;
  window.clearTimeout(statusFeedbackTimer);
  statusMessage.textContent = isError && text ? `${text}${/try again\.?$/i.test(text) ? "" : " Try again."}` : text;
  statusMessage.classList.toggle("error", isError);
  statusMessage.classList.toggle("status-subtle-success", subtle);
  statusMessage.classList.toggle("status-prominent-error", isError);
  statusMessage.hidden = !text;
  if (autoClear && text) {
    statusFeedbackTimer = window.setTimeout(() => {
      statusMessage.textContent = "";
      statusMessage.classList.remove("error", "status-subtle-success", "status-prominent-error");
      statusMessage.hidden = true;
    }, subtle ? 1400 : 2200);
  }
}

function applyStageUpdateToCandidate(candidate, stageKey, outcome, nextSteps) {
  if (!candidate) {
    return;
  }
  if (stageKey === "initial_screen") {
    candidate.screenOutcome = outcome;
    candidate.screenNextSteps = nextSteps;
    return;
  }
  if (stageKey === "first_interview") {
    candidate.firstInterviewOutcome = outcome;
    candidate.firstInterviewNextSteps = nextSteps;
    return;
  }
  if (stageKey === "second_interview") {
    candidate.secondInterviewOutcome = outcome;
    candidate.secondInterviewNextSteps = nextSteps;
  }
}

function setRecruitmentToolbarVisible(visible) {
  if (!recruitmentToolbarContent || !toggleRecruitmentToolbarBtn) {
    return;
  }
  recruitmentToolbarContent.hidden = !visible;
  toggleRecruitmentToolbarBtn.setAttribute("aria-expanded", visible ? "true" : "false");
  toggleRecruitmentToolbarBtn.setAttribute(
    "aria-label",
    visible ? "Hide search and import tools" : "Show search and import tools"
  );
  toggleRecruitmentToolbarBtn.setAttribute(
    "title",
    visible ? "Hide search and import tools" : "Show search and import tools"
  );
  toggleRecruitmentToolbarBtn.classList.toggle("is-open", visible);
}

function syncAddRecruitmentButton() {
  if (!addRecruitmentItemBtn) {
    return;
  }
  addRecruitmentItemBtn.disabled = createCandidateBusy;
}

function hasOneTouchLink(candidate) {
  return Boolean(cleanText(candidate?.oneTouchLink));
}

function canAddToOneTouch(candidate) {
  return ONE_TOUCH_ELIGIBLE_STATUSES.has(cleanText(candidate?.status));
}

function getInitialScreenUrl(candidateId) {
  const url = new URL("./initial-screen.html", window.location.href);
  url.searchParams.set("itemId", cleanText(candidateId));
  return url.toString();
}

function formatAppliedOnDate(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric) || numeric <= 0) {
    return "";
  }
  const parsed = new Date(numeric);
  return Number.isNaN(parsed.getTime()) ? "" : parsed.toLocaleDateString();
}

function normalizeJsonJobLocation(value) {
  const raw = cleanText(value);
  if (!raw) {
    return "";
  }
  return raw.replace(/\s*\([^)]*\)\s*$/, "").trim();
}

function extractRecruitmentJsonCandidate(payload) {
  const applicant = payload?.analytics?.applicant || payload?.applicant || {};
  const job = payload?.job || {};
  const screenerAnswers = Array.isArray(payload?.screenerQuestionsAndAnswers?.questionsAndAnswers)
    ? payload.screenerQuestionsAndAnswers.questionsAndAnswers
    : [];
  const indeedProfileUrl = cleanText(applicant?.publicProfileUrl || payload?.publicProfileUrl);

  const notes = [];
  const currentRole = [cleanText(applicant?.jobTitle), cleanText(applicant?.companyName)].filter(Boolean).join(" at ");
  if (currentRole) {
    notes.push(`Current role: ${currentRole}`);
  }
  const appliedOn = formatAppliedOnDate(payload?.appliedOnMillis);
  if (appliedOn) {
    notes.push(`Applied on: ${appliedOn}`);
  }
  if (cleanText(job?.jobTitle)) {
    notes.push(`Applied for: ${cleanText(job.jobTitle)}`);
  }

  const screenerSummary = screenerAnswers
    .slice(0, 6)
    .map((entry) => {
      const question = cleanText(entry?.question?.question);
      const answer =
        cleanText(entry?.answer?.label) ||
        cleanText(entry?.answer?.value) ||
        cleanText(entry?.answer);
      if (!question || !answer) {
        return "";
      }
      return `${question}: ${answer}`;
    })
    .filter(Boolean);
  if (screenerSummary.length) {
    notes.push(`Screener: ${screenerSummary.join(" | ")}`);
  }
  if (indeedProfileUrl) {
    notes.push(`Indeed Profile: ${indeedProfileUrl}`);
  }

  return {
    candidateName: cleanText(applicant?.fullName),
    status: "Organise Initial Call",
    email: cleanText(applicant?.email),
    phoneNumber: cleanText(applicant?.phoneNumber),
    livesIn: cleanText(applicant?.location?.city),
    location: normalizeJsonJobLocation(job?.jobLocation),
    source: "Indeed - JSON upload",
    indeedUrl: indeedProfileUrl,
    active: true,
    notes: notes.join("\n"),
  };
}

function applyRecruitmentJsonCandidate(prefill) {
  if (addCandidateNameInput) {
    addCandidateNameInput.value = cleanText(prefill?.candidateName);
  }
  if (addCandidateStatusSelect) {
    addCandidateStatusSelect.value = cleanText(prefill?.status) || "Organise Initial Call";
  }
  if (addCandidateEmailInput) {
    addCandidateEmailInput.value = cleanText(prefill?.email);
  }
  if (addCandidatePhoneInput) {
    addCandidatePhoneInput.value = cleanText(prefill?.phoneNumber);
  }
  if (addCandidateLivesInInput) {
    addCandidateLivesInInput.value = cleanText(prefill?.livesIn);
  }
  if (addCandidateJobLocationInput) {
    addCandidateJobLocationInput.value = cleanText(prefill?.location);
  }
  if (addCandidateSourceInput) {
    addCandidateSourceInput.value = cleanText(prefill?.source);
  }
  if (addCandidateIndeedUrlInput) {
    addCandidateIndeedUrlInput.value = cleanText(prefill?.indeedUrl);
  }
  if (addCandidateActiveInput) {
    addCandidateActiveInput.checked = prefill?.active !== false;
  }
  if (addCandidateUpdateExistingInput) {
    addCandidateUpdateExistingInput.checked = true;
  }
  if (addCandidateNotesInput) {
    addCandidateNotesInput.value = cleanText(prefill?.notes);
  }
}

async function handleAddRecruitmentJsonFile(file) {
  if (!file) {
    return;
  }
  if (addRecruitmentJsonFileName) {
    addRecruitmentJsonFileName.textContent = `Selected: ${file.name}`;
  }

  try {
    const raw = await file.text();
    const parsed = JSON.parse(raw);
    const prefill = extractRecruitmentJsonCandidate(parsed);
    if (!prefill.candidateName) {
      throw new Error("Could not find candidate details in this JSON file.");
    }
    applyRecruitmentJsonCandidate(prefill);
    setAddRecruitmentError("");
  } catch (error) {
    console.error(error);
    setAddRecruitmentError(error?.message || "Could not read candidate JSON.");
  }
}

function setAddRecruitmentError(message = "") {
  if (!addRecruitmentError) {
    return;
  }
  const text = cleanText(message);
  addRecruitmentError.hidden = !text;
  addRecruitmentError.textContent = text;
}

function renderAddRecruitmentStatusOptions() {
  if (!addCandidateStatusSelect) {
    return;
  }
  const current = cleanText(addCandidateStatusSelect.value);
  addCandidateStatusSelect.innerHTML = "";
  for (const status of recruitmentStatusOptions) {
    const option = document.createElement("option");
    option.value = status;
    option.textContent = status;
    addCandidateStatusSelect.appendChild(option);
  }
  const fallback = recruitmentStatusOptions.includes("Organise Initial Call")
    ? "Organise Initial Call"
    : recruitmentStatusOptions[0] || "";
  addCandidateStatusSelect.value = recruitmentStatusOptions.includes(current) ? current : fallback;
}

function resetAddRecruitmentForm() {
  addRecruitmentForm?.reset();
  renderAddRecruitmentStatusOptions();
  if (addCandidateActiveInput) {
    addCandidateActiveInput.checked = true;
  }
  if (addCandidateUpdateExistingInput) {
    addCandidateUpdateExistingInput.checked = true;
  }
  if (addRecruitmentJsonFileName) {
    addRecruitmentJsonFileName.textContent = "No JSON file selected.";
  }
  if (addRecruitmentJsonFileInput) {
    addRecruitmentJsonFileInput.value = "";
  }
  setAddRecruitmentError("");
}

function setCreateCandidateBusy(disabled) {
  createCandidateBusy = disabled;
  syncAddRecruitmentButton();
  if (saveRecruitmentCandidateBtn) {
    saveRecruitmentCandidateBtn.disabled = disabled;
  }
  if (cancelAddRecruitmentBtn) {
    cancelAddRecruitmentBtn.disabled = disabled;
  }
  for (const field of [
    addRecruitmentJsonFileInput,
    addCandidateNameInput,
    addCandidateStatusSelect,
    addCandidateEmailInput,
    addCandidatePhoneInput,
    addCandidateLivesInInput,
    addCandidateJobLocationInput,
    addCandidateSourceInput,
    addCandidateIndeedUrlInput,
    addCandidateActiveInput,
    addCandidateUpdateExistingInput,
    addCandidateNotesInput,
  ]) {
    if (field) {
      field.disabled = disabled;
    }
  }
  if (addRecruitmentJsonDropZone) {
    addRecruitmentJsonDropZone.classList.toggle("is-disabled", disabled);
  }
}

function openAddRecruitmentModal() {
  resetAddRecruitmentForm();
  if (addRecruitmentModal) {
    addRecruitmentModal.hidden = false;
  }
  addCandidateNameInput?.focus();
}

function closeAddRecruitmentModal(options = {}) {
  if (createCandidateBusy && options.force !== true) {
    return;
  }
  if (addRecruitmentModal) {
    addRecruitmentModal.hidden = true;
  }
  setAddRecruitmentError("");
}

function openCandidateDetail(candidate) {
  if (!candidate?.id || !candidateDetailModal) {
    return;
  }
  setDetail(candidate);
  candidateDetailModal.hidden = false;
}

function closeCandidateDetail() {
  if (candidateDetailModal) {
    candidateDetailModal.hidden = true;
  }
}

function setDetailFormEnabled(enabled) {
  for (const field of Object.values(detailInputs)) {
    if (!field) {
      continue;
    }
    field.disabled = !enabled || detailSaveBusy;
  }
  if (saveDetailBtn) {
    saveDetailBtn.disabled = !enabled || detailSaveBusy;
  }
}

function normalizePhoneForActions(phoneNumber) {
  const raw = cleanText(phoneNumber);
  if (!raw) {
    return "";
  }
  let digits = raw.replace(/[^\d+]/g, "");
  if (digits.startsWith("00")) {
    digits = `+${digits.slice(2)}`;
  }
  if (!digits.startsWith("+")) {
    const numeric = digits.replace(/\D/g, "");
    if (numeric.startsWith("0")) {
      digits = `+44${numeric.slice(1)}`;
    } else if (numeric.startsWith("44")) {
      digits = `+${numeric}`;
    } else {
      digits = `+${numeric}`;
    }
  }
  return digits.replace(/(?!^\+)\D/g, "");
}

function getWhatsAppMessage(candidateName) {
  const firstName = cleanText(candidateName).split(/\s+/)[0] || "there";
  return `Hi ${firstName}, thanks for applying to Thrive Homecare. Are you available for a quick initial chat? Chris`;
}

function getWhatsAppUrl(phoneNumber, candidateName) {
  const normalized = normalizePhoneForActions(phoneNumber).replace(/\D/g, "");
  if (!normalized) {
    return "";
  }
  const url = new URL("https://web.whatsapp.com/send");
  url.searchParams.set("phone", normalized);
  url.searchParams.set("text", getWhatsAppMessage(candidateName));
  return url.toString();
}

function getTeamsCallUrl(phoneNumber) {
  const normalized = normalizePhoneForActions(phoneNumber);
  if (!normalized) {
    return "";
  }
  return `https://teams.microsoft.com/l/call/0/0?users=${encodeURIComponent(normalized)}`;
}

function getIndeedProfileUrl(candidate) {
  return cleanText(candidate?.indeedProfileUrl);
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

function setSelectLoading(selectEl, placeholder) {
  if (!selectEl) {
    return;
  }
  selectEl.disabled = true;
  selectEl.innerHTML = "";
  const option = document.createElement("option");
  option.value = "";
  option.textContent = placeholder;
  selectEl.appendChild(option);
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
    positions: Array.isArray(options?.positions) ? options.positions : [],
    statuses: Array.isArray(options?.statuses) ? options.statuses : [],
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

  setSelectLoading(oneTouchAreaSelect, "Loading areas...");
  setSelectLoading(oneTouchRecruitmentSourceSelect, "Loading recruitment sources...");
  setSelectLoading(oneTouchPositionSelect, "Loading positions...");
  setSelectLoading(oneTouchStatusSelect, "Loading statuses...");

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
      oneTouchAreaSelect.value =
        pickBestOption(options.areas, ONE_TOUCH_DEFAULT_AREA) ||
        pickBestOption(options.areas, candidate?.earmarkedFor) ||
        "";
      oneTouchAreaSelect.disabled = false;
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
      oneTouchRecruitmentSourceSelect.disabled = false;
    }
    if (oneTouchPositionSelect) {
      oneTouchPositionSelect.innerHTML = '<option value="">Select position</option>';
      for (const position of options.positions) {
        const option = document.createElement("option");
        option.value = position;
        option.textContent = position;
        oneTouchPositionSelect.appendChild(option);
      }
      oneTouchPositionSelect.value =
        pickBestOption(options.positions, ONE_TOUCH_DEFAULT_POSITION) ||
        pickBestOption(options.positions, "Carer") ||
        "";
      oneTouchPositionSelect.disabled = false;
    }
    if (oneTouchStatusSelect) {
      oneTouchStatusSelect.innerHTML = '<option value="">Select status</option>';
      for (const status of options.statuses) {
        const option = document.createElement("option");
        option.value = status;
        option.textContent = status;
        oneTouchStatusSelect.appendChild(option);
      }
      oneTouchStatusSelect.value =
        pickBestOption(options.statuses, ONE_TOUCH_DEFAULT_STATUS) ||
        pickBestOption(options.statuses, candidate?.status) ||
        "";
      oneTouchStatusSelect.disabled = false;
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
    return "Organise Initial Call";
  }
  if (/\b(contacting|applied|application|new)\b/.test(merged)) {
    return "Organise Initial Call";
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
  return status || "Organise Initial Call";
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

function toSortTimestamp(value) {
  const raw = cleanText(value);
  if (!raw) {
    return 0;
  }
  const parsed = new Date(raw);
  return Number.isNaN(parsed.getTime()) ? 0 : parsed.getTime();
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

function setOneTouchButton(url) {
  if (!openOneTouchBtn) {
    return;
  }
  const cleanUrl = cleanText(url);
  const enabled = Boolean(cleanUrl);
  openOneTouchBtn.href = enabled ? cleanUrl : "#";
  openOneTouchBtn.setAttribute("aria-disabled", enabled ? "false" : "true");
  openOneTouchBtn.classList.toggle("is-disabled", !enabled);
}

function setIndeedButton(url) {
  if (!openIndeedBtn) {
    return;
  }
  const cleanUrl = cleanText(url);
  const enabled = Boolean(cleanUrl);
  openIndeedBtn.href = enabled ? cleanUrl : "#";
  openIndeedBtn.setAttribute("aria-disabled", enabled ? "false" : "true");
  openIndeedBtn.classList.toggle("is-disabled", !enabled);
}

function syncActiveToggleButton(button, isActive) {
  if (!button) {
    return;
  }
  button.classList.toggle("is-active", isActive === true);
  button.setAttribute("aria-pressed", isActive === true ? "true" : "false");
  const label = button.querySelector(".recruitment-active-toggle-label");
  if (label) {
    label.textContent = isActive === true ? "Active" : "Inactive";
  }
}

function candidateMatchesActiveFilter(candidate) {
  const candidateId = cleanText(candidate?.id);
  if (dismissedInactiveReviewIds.has(candidateId)) {
    return false;
  }
  const selectedActive = cleanText(activeFilterSelect?.value || "active");
  if (selectedActive === "active") {
    return candidate?.active === true || pendingInactiveReviewIds.has(candidateId);
  }
  if (selectedActive === "inactive") {
    return candidate?.active === false;
  }
  return true;
}

function isInactiveReviewPending(candidate) {
  return pendingInactiveReviewIds.has(cleanText(candidate?.id));
}

function renderStatusUpdateOptions() {
  if (!statusUpdateSelect) {
    return;
  }
  const current = cleanText(statusUpdateSelect.value);
  statusUpdateSelect.innerHTML = '<option value="">Select status</option>';
  for (const status of recruitmentStatusOptions) {
    const option = document.createElement("option");
    option.value = status;
    option.textContent = status;
    statusUpdateSelect.appendChild(option);
  }
  statusUpdateSelect.value = recruitmentStatusOptions.includes(current) ? current : "";
}

function closeStatusQuickMenu() {
  statusQuickMenuCandidateId = "";
  if (statusQuickMenu) {
    statusQuickMenu.hidden = true;
    statusQuickMenu.style.left = "";
    statusQuickMenu.style.top = "";
  }
  if (statusQuickMenuList) {
    statusQuickMenuList.innerHTML = "";
  }
}

function openStatusQuickMenu(candidate, anchorEl) {
  if (!statusQuickMenu || !statusQuickMenuList || !anchorEl || !candidate?.id) {
    return;
  }
  const candidateId = cleanText(candidate.id);
  if (!candidateId) {
    return;
  }

  statusQuickMenuCandidateId = candidateId;
  const currentStatus = cleanText(candidate.status);
  const optionsHtml = recruitmentStatusOptions.map((status) => {
    const activeClass = currentStatus === status ? " is-active" : "";
    return `<button type="button" class="status-quick-option${activeClass}" data-status="${escapeHtml(status)}">${escapeHtml(status)}</button>`;
  }).join("");
  statusQuickMenuList.innerHTML = optionsHtml;

  for (const btn of statusQuickMenuList.querySelectorAll(".status-quick-option")) {
    btn.addEventListener("click", async (event) => {
      event.stopPropagation();
      if (statusUpdateBusy) {
        return;
      }
      const nextStatus = cleanText(btn.getAttribute("data-status"));
      if (!nextStatus) {
        return;
      }
      statusUpdateBusy = true;
      try {
        await updateCandidateStatusById(candidateId, nextStatus);
        closeStatusQuickMenu();
      } catch (error) {
        console.error(error);
        setStatus(error?.message || "Could not update status.", true);
      } finally {
        statusUpdateBusy = false;
      }
    });
  }

  const rect = anchorEl.getBoundingClientRect();
  statusQuickMenu.hidden = false;
  statusQuickMenu.style.visibility = "hidden";

  const viewportPad = 12;
  const offset = 8;
  const menuWidth = statusQuickMenu.offsetWidth || 240;
  const menuHeight = statusQuickMenu.offsetHeight || 280;
  const viewportLeft = window.scrollX;
  const viewportTop = window.scrollY;
  const viewportRight = viewportLeft + window.innerWidth;
  const viewportBottom = viewportTop + window.innerHeight;

  let left = rect.left + window.scrollX;
  if (left + menuWidth + viewportPad > viewportRight) {
    left = viewportRight - menuWidth - viewportPad;
  }
  left = Math.max(viewportLeft + viewportPad, left);

  const topBelow = rect.bottom + window.scrollY + offset;
  const topAbove = rect.top + window.scrollY - menuHeight - offset;
  const top =
    topBelow + menuHeight + viewportPad <= viewportBottom || topAbove < viewportTop + viewportPad
      ? topBelow
      : topAbove;

  statusQuickMenu.style.top = `${top}px`;
  statusQuickMenu.style.left = `${left}px`;
  statusQuickMenu.style.visibility = "";
}

function setDetail(candidate) {
  if (!candidate) {
    if (candidateDetailModalTitle) {
      candidateDetailModalTitle.textContent = "Candidate Detail";
    }
    detailFields.candidateName.textContent = "Select a candidate";
    detailFields.location.textContent = "-";
    detailFields.status.textContent = "-";
    detailFields.active.textContent = "-";
    detailFields.source.textContent = "-";
    detailFields.phoneNumber.textContent = "-";
    detailFields.email.textContent = "-";
    renderTagPreview(detailFields.tags, []);
    detailFields.interviewBooked.textContent = "-";
    detailFields.interviewWith.textContent = "-";
    detailFields.keepInMind.textContent = "-";
    detailFields.livesIn.textContent = "-";
    detailFields.firstInterviewDate.textContent = "-";
    detailFields.earmarkedFor.textContent = "-";
    detailFields.created.textContent = "-";
    detailFields.oneTouchLink.textContent = "-";
    detailFields.notes.textContent = "-";
    if (detailInputs.candidateName) {
      detailInputs.candidateName.value = "";
      detailInputs.location.value = "";
      detailInputs.source.value = "";
      detailInputs.phoneNumber.value = "";
      detailInputs.email.value = "";
      detailInputs.indeedUrl.value = "";
      detailInputs.livesIn.value = "";
      detailInputs.earmarkedFor.value = "";
      detailInputs.keepInMind.checked = false;
      detailInputs.tags.value = "";
      detailInputs.notes.value = "";
    }
    renderTagPreview(detailTagsPreview, []);
    setDetailFormEnabled(false);
    setIndeedButton("");
    setOneTouchButton("");
    if (statusUpdateSelect) {
      statusUpdateSelect.value = "";
      statusUpdateSelect.disabled = true;
    }
    if (saveStatusBtn) {
      saveStatusBtn.disabled = true;
    }
    return;
  }

  if (candidateDetailModalTitle) {
    candidateDetailModalTitle.textContent = cleanText(candidate.candidateName) || "Candidate Detail";
  }
  detailFields.candidateName.textContent = cleanText(candidate.candidateName) || "-";
  detailFields.location.textContent = cleanText(candidate.location) || "-";
  detailFields.status.textContent = cleanText(candidate.status) || "-";
  detailFields.active.textContent = formatBoolean(candidate.active);
  detailFields.source.textContent = cleanText(candidate.source) || "-";
  detailFields.phoneNumber.textContent = cleanText(candidate.phoneNumber) || "-";
  detailFields.email.textContent = cleanText(candidate.email) || "-";
  renderTagPreview(detailFields.tags, candidate.tags);
  detailFields.interviewBooked.textContent = formatBoolean(candidate.interviewBooked);
  detailFields.interviewWith.textContent = cleanText(candidate.interviewWith) || "-";
  detailFields.keepInMind.textContent = formatBoolean(candidate.keepInMind);
  detailFields.livesIn.textContent = cleanText(candidate.livesIn) || "-";
  detailFields.firstInterviewDate.textContent = formatDate(candidate.firstInterviewDate);
  detailFields.earmarkedFor.textContent = cleanText(candidate.earmarkedFor) || "-";
  detailFields.created.textContent = formatDate(candidate.created);
  setIndeedButton(candidate.indeedProfileUrl);
  setLinkField(detailFields.oneTouchLink, candidate.oneTouchLink);
  setOneTouchButton(candidate.oneTouchLink);
  detailFields.notes.textContent = cleanText(candidate.notes) || "-";
  if (detailInputs.candidateName) {
    detailInputs.candidateName.value = cleanText(candidate.candidateName);
    detailInputs.location.value = cleanText(candidate.location);
    detailInputs.source.value = cleanText(candidate.source);
    detailInputs.phoneNumber.value = cleanText(candidate.phoneNumber);
    detailInputs.email.value = cleanText(candidate.email);
    detailInputs.indeedUrl.value = cleanText(candidate.indeedProfileUrl);
    detailInputs.livesIn.value = cleanText(candidate.livesIn);
    detailInputs.earmarkedFor.value = cleanText(candidate.earmarkedFor);
    detailInputs.keepInMind.checked = candidate.keepInMind === true;
    detailInputs.tags.value = normalizeTagString(candidate.tags);
    detailInputs.notes.value = cleanText(candidate.notes);
  }
  renderTagPreview(detailTagsPreview, candidate.tags);
  setDetailFormEnabled(true);
  renderStatusUpdateOptions();
  if (statusUpdateSelect) {
    statusUpdateSelect.value = cleanText(candidate.status);
    statusUpdateSelect.disabled = false;
  }
  if (saveStatusBtn) {
    saveStatusBtn.disabled = statusUpdateBusy;
  }
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
  const ownerOptions = Array.from(
    new Set(
      [...recruitmentOwnerOptions.map((option) => cleanText(option.label)), ...allCandidates.map((candidate) => cleanText(candidate.currentOwner))]
        .filter(Boolean)
    )
  ).sort((a, b) => a.localeCompare(b));

  const selectedLocation = cleanText(locationFilterSelect.value || "all");
  const selectedStatus = cleanText(statusFilterSelect.value || STATUS_FILTER_DEFAULT);
  const selectedSource = cleanText(sourceFilterSelect.value || "all");
  const selectedOwner = cleanText(ownerFilterSelect?.value || "all");
  const selectedActive = cleanText(activeFilterSelect?.value || "active");
  const selectedSort = cleanText(sortFilterSelect?.value || "updated_desc");

  locationFilterSelect.innerHTML = '<option value="all">All locations</option>';
  statusFilterSelect.innerHTML =
    `<option value="${STATUS_FILTER_DEFAULT}">All except rejected</option><option value="all">All statuses</option>`;
  sourceFilterSelect.innerHTML = '<option value="all">All sources</option>';
  if (ownerFilterSelect) {
    ownerFilterSelect.innerHTML = '<option value="all">All owners</option>';
  }

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
  for (const owner of ownerOptions) {
    const option = document.createElement("option");
    option.value = owner;
    option.textContent = owner;
    ownerFilterSelect?.appendChild(option);
  }

  locationFilterSelect.value = locationOptions.includes(selectedLocation) ? selectedLocation : "all";
  statusFilterSelect.value =
    selectedStatus === STATUS_FILTER_DEFAULT || statusOptions.includes(selectedStatus) ? selectedStatus : STATUS_FILTER_DEFAULT;
  sourceFilterSelect.value = sourceOptions.includes(selectedSource) ? selectedSource : "all";
  if (ownerFilterSelect) {
    ownerFilterSelect.value = ownerOptions.includes(selectedOwner) ? selectedOwner : "all";
  }
  if (activeFilterSelect) {
    activeFilterSelect.value = ["active", "inactive", "all"].includes(selectedActive) ? selectedActive : "active";
  }
  if (sortFilterSelect) {
    sortFilterSelect.value = [
      "updated_desc",
      "updated_asc",
      "created_desc",
      "created_asc",
      "name_asc",
      "name_desc",
      "location_asc",
      "location_desc",
      "status_asc",
      "status_desc",
    ].includes(selectedSort)
      ? selectedSort
      : "updated_desc";
  }
  syncSortHeaderButtons();
}

function getFilteredCandidates() {
  const selectedSort = cleanText(sortFilterSelect?.value || "updated_desc");
  const baseFiltered = getBaseFilteredCandidates();
  const filtered = baseFiltered.filter((candidate) => candidateMatchesStageMode(candidate));

  filtered.sort((left, right) => {
    if (selectedSort === "updated_asc") {
      return toSortTimestamp(left.updated) - toSortTimestamp(right.updated);
    }
    if (selectedSort === "created_desc") {
      return toSortTimestamp(right.created) - toSortTimestamp(left.created);
    }
    if (selectedSort === "created_asc") {
      return toSortTimestamp(left.created) - toSortTimestamp(right.created);
    }
    if (selectedSort === "name_asc") {
      return getFirstNameSortValue(left.candidateName).localeCompare(getFirstNameSortValue(right.candidateName));
    }
    if (selectedSort === "name_desc") {
      return getFirstNameSortValue(right.candidateName).localeCompare(getFirstNameSortValue(left.candidateName));
    }
    if (selectedSort === "location_asc") {
      return cleanText(left.location).localeCompare(cleanText(right.location));
    }
    if (selectedSort === "location_desc") {
      return cleanText(right.location).localeCompare(cleanText(left.location));
    }
    if (selectedSort === "status_asc") {
      return cleanText(left.status).localeCompare(cleanText(right.status));
    }
    if (selectedSort === "status_desc") {
      return cleanText(right.status).localeCompare(cleanText(left.status));
    }
    return toSortTimestamp(right.updated) - toSortTimestamp(left.updated);
  });

  return filtered;
}

function renderCandidates() {
  const renderToken = ++stageModeCountRenderToken;
  const filtered = getFilteredCandidates();
  recruitmentTableBody.innerHTML = "";
  closeStatusQuickMenu();

  window.requestAnimationFrame(() => {
    if (renderToken !== stageModeCountRenderToken) {
      return;
    }
    updateStageModeFilterButtons(getBaseFilteredCandidates());
  });

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
    const whatsappUrl = getWhatsAppUrl(candidate.phoneNumber, candidate.candidateName);
    const teamsCallUrl = getTeamsCallUrl(candidate.phoneNumber);
    const indeedUrl = getIndeedProfileUrl(candidate);
    const showInactiveReview = isInactiveReviewPending(candidate);
    tr.innerHTML = `
      <td>${escapeHtml(cleanText(candidate.candidateName) || "-")}</td>
      <td>${escapeHtml(cleanText(candidate.location) || "-")}</td>
      <td>
        <div class="recruitment-status-cell${showInactiveReview ? " is-inactive-review" : ""}">
          <button type="button" class="status-pill-trigger">${escapeHtml(cleanText(candidate.status) || "-")}</button>
          <label class="recruitment-owner-field">
            <span class="recruitment-owner-label">Current owner</span>
            <select class="recruitment-owner-select">
              <option value="">Unassigned</option>
              ${recruitmentOwnerOptions.map(
                (option) =>
                  `<option value="${escapeHtml(option.label)}"${getCurrentOwnerLabel(candidate) === option.label ? " selected" : ""}>${escapeHtml(
                    option.label
                  )}</option>`
              ).join("")}
            </select>
          </label>
          <button
            type="button"
            class="recruitment-active-toggle${candidate.active ? " is-active" : ""}"
            aria-pressed="${candidate.active ? "true" : "false"}"
            ${activeUpdateBusy ? " disabled" : ""}
          >
            <span class="recruitment-active-toggle-track">
            <span class="recruitment-active-toggle-thumb"></span>
            </span>
            <span class="recruitment-active-toggle-label">${candidate.active ? "Active" : "Inactive"}</span>
          </button>
          ${
            showInactiveReview
              ? `<div class="recruitment-inactive-review">
                  <label class="recruitment-inactive-review-checkbox">
                    <input type="checkbox" class="recruitment-keep-in-mind-toggle"${candidate.keepInMind ? " checked" : ""} />
                    <span>Keep in mind</span>
                  </label>
                  <button type="button" class="secondary recruitment-inactive-review-done">Done</button>
                </div>`
              : ""
          }
        </div>
      </td>
      <td>${renderStageSummary(candidate)}</td>
      <td>
        <div class="recruitment-action-stack">
          <button type="button" class="secondary recruitment-screen-link recruitment-detail-trigger">View details</button>
          <a class="secondary recruitment-screen-link" href="${escapeHtml(getInitialScreenUrl(candidate.id))}">Initial Screen</a>
          ${
            whatsappUrl
              ? `<a class="recruitment-quick-link recruitment-whatsapp-link" href="${escapeHtml(whatsappUrl)}" target="_blank" rel="noopener noreferrer">Text in WhatsApp</a>`
              : ""
          }
          ${
            teamsCallUrl
              ? `<a class="recruitment-quick-link recruitment-teams-link" href="${escapeHtml(teamsCallUrl)}" target="_blank" rel="noopener noreferrer">Call in Teams</a>`
              : ""
          }
          ${
            indeedUrl
              ? `<a class="recruitment-quick-link" href="${escapeHtml(indeedUrl)}" target="_blank" rel="noopener noreferrer">Open in Indeed</a>`
              : ""
          }
          ${
            hasOneTouchLink(candidate)
              ? `<a
                  class="button-link-one-touch recruitment-open-link"
                  href="${escapeHtml(cleanText(candidate.oneTouchLink))}"
                  target="_blank"
                  rel="noopener noreferrer"
                  >Open in OneTouch</a
                >`
              : canAddToOneTouch(candidate)
                ? `<button type="button" class="secondary recruitment-add-btn"${addToOneTouchBusy ? " disabled" : ""}>Add to OneTouch</button>`
                : ""
          }
        </div>
      </td>
    `;

    tr.addEventListener("click", () => {
      selectedCandidateId = candidate.id;
      setDetail(candidate);
      renderCandidates();
      openCandidateDetail(candidate);
    });

    const detailBtn = tr.querySelector(".recruitment-detail-trigger");
    detailBtn?.addEventListener("click", (event) => {
      event.stopPropagation();
      selectedCandidateId = candidate.id;
      setDetail(candidate);
      renderCandidates();
      openCandidateDetail(candidate);
    });
    const addBtn = tr.querySelector(".recruitment-add-btn");
    addBtn?.addEventListener("click", async (event) => {
      event.stopPropagation();
      if (addToOneTouchBusy) {
        return;
      }
      await openOneTouchPicker(candidate);
    });
    const statusTrigger = tr.querySelector(".status-pill-trigger");
    statusTrigger?.addEventListener("click", (event) => {
      event.stopPropagation();
      selectedCandidateId = candidate.id;
      setDetail(candidate);
      if (statusUpdateBusy || addToOneTouchBusy) {
        return;
      }
      openStatusQuickMenu(candidate, statusTrigger);
    });
    const ownerSelect = tr.querySelector(".recruitment-owner-select");
    ownerSelect?.addEventListener("click", (event) => {
      event.stopPropagation();
    });
    ownerSelect?.addEventListener("change", async (event) => {
      event.stopPropagation();
      if (ownerUpdateBusy) {
        return;
      }
      const nextOwner = cleanText(ownerSelect.value);
      const previousOwner = cleanText(candidate.currentOwner);
      if (nextOwner === previousOwner) {
        return;
      }
      ownerUpdateBusy = true;
      ownerSelect.disabled = true;
      try {
        await updateCandidateOwnerById(candidate.id, nextOwner);
      } catch (error) {
        console.error(error);
        ownerSelect.value = previousOwner;
        setStatus(error?.message || "Could not update current owner.", true, { autoClear: false });
      } finally {
        ownerUpdateBusy = false;
        if (ownerSelect && document.body.contains(ownerSelect)) {
          ownerSelect.disabled = false;
        }
        renderCandidates();
      }
    });
    const keepInMindToggle = tr.querySelector(".recruitment-keep-in-mind-toggle");
    keepInMindToggle?.addEventListener("click", (event) => {
      event.stopPropagation();
    });
    keepInMindToggle?.addEventListener("change", async (event) => {
      event.stopPropagation();
      const nextKeepInMind = keepInMindToggle.checked === true;
      const previousKeepInMind = candidate.keepInMind === true;
      candidate.keepInMind = nextKeepInMind;
      renderCandidates();
      try {
        await directoryApi.updateRecruitmentKeepInMind({
          itemId: candidate.id,
          keepInMind: nextKeepInMind,
        });
        setStatus(`Keep in mind ${nextKeepInMind ? "enabled" : "cleared"}.`, false, { subtle: true });
      } catch (error) {
        candidate.keepInMind = previousKeepInMind;
        renderCandidates();
        console.error(error);
        setStatus(error?.message || "Could not update keep in mind.", true, { autoClear: false });
      }
    });
    const inactiveReviewDoneBtn = tr.querySelector(".recruitment-inactive-review-done");
    inactiveReviewDoneBtn?.addEventListener("click", (event) => {
      event.stopPropagation();
      const candidateId = cleanText(candidate.id);
      pendingInactiveReviewIds.delete(candidateId);
      dismissedInactiveReviewIds.add(candidateId);
      renderCandidates();
    });
    const activeToggle = tr.querySelector(".recruitment-active-toggle");
    activeToggle?.addEventListener("click", async (event) => {
      event.stopPropagation();
      if (activeUpdateBusy) {
        return;
      }
      const nextActive = candidate.active !== true;
      const candidateId = cleanText(candidate.id);
      activeUpdateBusy = true;
      if (activeToggle) {
        activeToggle.disabled = true;
      }
      try {
        if (nextActive === false) {
          pendingInactiveReviewIds.add(candidateId);
          dismissedInactiveReviewIds.delete(candidateId);
        } else {
          pendingInactiveReviewIds.delete(candidateId);
          dismissedInactiveReviewIds.delete(candidateId);
        }
        await updateCandidateActiveById(candidate.id, nextActive);
        syncActiveToggleButton(activeToggle, nextActive);
        if (candidate.id === selectedCandidateId) {
          setDetail(candidate);
        }
        window.clearTimeout(activeHideRefreshTimer);
        renderCandidates();
      } catch (error) {
        if (nextActive === false) {
          pendingInactiveReviewIds.delete(candidateId);
        } else if (candidate.active === false) {
          pendingInactiveReviewIds.add(candidateId);
        }
        console.error(error);
        setStatus(error?.message || "Could not update active status.", true, { autoClear: false });
      } finally {
        activeUpdateBusy = false;
        if (activeToggle && document.body.contains(activeToggle)) {
          activeToggle.disabled = false;
        }
      }
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

async function saveRecruitmentStage(itemId, stageKey, outcome, nextSteps) {
  const cleanItemId = cleanText(itemId);
  const cleanStageKey = cleanText(stageKey);
  if (!cleanItemId || !cleanStageKey) {
    return;
  }
  const candidate = allCandidates.find((item) => cleanText(item.id) === cleanItemId);
  const previousValues = candidate
    ? {
        screenOutcome: candidate.screenOutcome,
        screenNextSteps: candidate.screenNextSteps,
        firstInterviewOutcome: candidate.firstInterviewOutcome,
        firstInterviewNextSteps: candidate.firstInterviewNextSteps,
        secondInterviewOutcome: candidate.secondInterviewOutcome,
        secondInterviewNextSteps: candidate.secondInterviewNextSteps,
      }
    : null;
  applyStageUpdateToCandidate(candidate, cleanStageKey, cleanText(outcome), cleanText(nextSteps));
  openStageKeys.add(`${cleanItemId}:${cleanStageKey}`);
  stageUpdateBusyKey = `${cleanItemId}:${cleanStageKey}`;
  renderCandidates();
  try {
    await directoryApi.updateRecruitmentStage({
      itemId: cleanItemId,
      stageKey: cleanStageKey,
      outcome: cleanText(outcome),
      nextSteps: cleanText(nextSteps),
    });
    renderCandidates();
    setStatus("Stage notes saved.", false, { subtle: true });
  } catch (error) {
    console.error(error);
    if (candidate && previousValues) {
      Object.assign(candidate, previousValues);
    }
    stageUpdateBusyKey = "";
    renderCandidates();
    setStatus(error?.message || "Could not save stage notes.", true, { autoClear: false });
    return;
  }
  stageUpdateBusyKey = "";
  renderCandidates();
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
  const selectedPosition = cleanText(oneTouchPositionSelect?.value);
  const selectedStatus = cleanText(oneTouchStatusSelect?.value);
  if (!selectedArea || !selectedRecruitmentSource || !selectedPosition || !selectedStatus) {
    setOneTouchPickerError("Select area, recruitment source, position, and status.");
    return;
  }

  setAddButtonsBusy(true);
  try {
    const result = await directoryApi.addRecruitmentCandidateToOneTouch({
      itemId: cleanItemId,
      area: selectedArea,
      recruitmentSource: selectedRecruitmentSource,
      position: selectedPosition,
      status: selectedStatus,
    });
    if (result?.item) {
      upsertCandidateInCache(result.item);
    } else if (result?.itemId && result?.oneTouchLink) {
      const existing = allCandidates.find((candidate) => candidate.id === cleanText(result.itemId));
      if (existing) {
        upsertCandidateInCache({
          ...existing,
          oneTouchLink: cleanText(result.oneTouchLink),
        });
      }
    }
    renderCandidates();
    closeOneTouchPicker();
    setStatus(`Candidate added to OneTouch (ID: ${cleanText(result?.oneTouchId) || "-"})`, false, { subtle: true });
  } catch (error) {
    console.error(error);
    setOneTouchPickerError(error?.message || "Could not add candidate to OneTouch.");
    setStatus(error?.message || "Could not add candidate to OneTouch.", true, { autoClear: false });
  } finally {
    setAddButtonsBusy(false);
  }
}

async function updateCandidateStatusById(candidateId, selectedStatus) {
  const targetId = cleanText(candidateId);
  const nextStatus = cleanText(selectedStatus);
  if (!targetId || !nextStatus) {
    setStatus("Select a status first.", true, { autoClear: false });
    return false;
  }
  const candidate = allCandidates.find((item) => item.id === targetId);
  const previousStatus = candidate ? cleanText(candidate.status) : "";
  if (candidate) {
    candidate.status = nextStatus;
  }
  renderFilterOptions();
  renderCandidates();

  try {
    await directoryApi.updateRecruitmentStatus({
      itemId: targetId,
      status: nextStatus,
    });
  } catch (error) {
    if (candidate) {
      candidate.status = previousStatus;
    }
    renderFilterOptions();
    renderCandidates();
    throw error;
  }

  if (targetId === selectedCandidateId && statusUpdateSelect) {
    statusUpdateSelect.value = nextStatus;
  }

  renderFilterOptions();
  renderCandidates();
  setStatus(`Status updated to ${nextStatus}.`, false, { subtle: true });
  return true;
}

async function updateCandidateActiveById(candidateId, nextActive) {
  const targetId = cleanText(candidateId);
  if (!targetId || typeof nextActive !== "boolean") {
    setStatus("Could not update active status.", true, { autoClear: false });
    return false;
  }
  const candidate = allCandidates.find((item) => item.id === targetId);
  const previousActive = candidate?.active;
  if (candidate) {
    candidate.active = nextActive;
  }
  renderFilterOptions();
  renderCandidates();

  try {
    await directoryApi.updateRecruitmentActive({
      itemId: targetId,
      active: nextActive,
    });
  } catch (error) {
    if (candidate) {
      candidate.active = previousActive;
    }
    renderFilterOptions();
    renderCandidates();
    throw error;
  }

  renderFilterOptions();
  renderCandidates();
  setStatus(`Candidate marked ${nextActive ? "active" : "inactive"}.`, false, { subtle: true });
  return true;
}

async function updateCandidateOwnerById(candidateId, nextOwnerEmail) {
  const targetId = cleanText(candidateId);
  if (!targetId) {
    setStatus("Could not update current owner.", true, { autoClear: false });
    return false;
  }

  const candidate = allCandidates.find((item) => item.id === targetId);
  const previousOwner = cleanText(candidate?.currentOwner);
  const nextOwner = recruitmentOwnerOptions.find((option) => cleanText(option.label) === cleanText(nextOwnerEmail)) || null;
  if (candidate) {
    candidate.currentOwner = nextOwner?.label || "";
  }
  renderFilterOptions();
  renderCandidates();

  try {
    await directoryApi.updateRecruitmentOwner({
      itemId: targetId,
      currentOwner: nextOwner?.label || "",
    });
  } catch (error) {
    if (candidate) {
      candidate.currentOwner = previousOwner;
    }
    renderFilterOptions();
    renderCandidates();
    throw error;
  }

  renderFilterOptions();
  renderCandidates();
  setStatus(`Current owner updated to ${nextOwner?.label || "Unassigned"}.`, false, { subtle: true });
  return true;
}

async function saveCandidateDetails() {
  if (detailSaveBusy || !selectedCandidateId) {
    return;
  }

  const payload = {
    itemId: selectedCandidateId,
    candidateName: cleanText(detailInputs.candidateName?.value),
    location: cleanText(detailInputs.location?.value),
    source: cleanText(detailInputs.source?.value),
    phoneNumber: cleanText(detailInputs.phoneNumber?.value),
    email: cleanText(detailInputs.email?.value),
    indeedUrl: cleanText(detailInputs.indeedUrl?.value),
    livesIn: cleanText(detailInputs.livesIn?.value),
    earmarkedFor: cleanText(detailInputs.earmarkedFor?.value),
    keepInMind: detailInputs.keepInMind?.checked === true,
    tags: normalizeTagString(detailInputs.tags?.value),
    notes: cleanText(detailInputs.notes?.value),
  };

  detailSaveBusy = true;
  setDetailFormEnabled(false);

  try {
    await directoryApi.updateRecruitmentDetails(payload);
    const candidate = allCandidates.find((item) => item.id === selectedCandidateId);
    if (candidate) {
      candidate.candidateName = payload.candidateName;
      candidate.location = payload.location;
      candidate.source = payload.source;
      candidate.phoneNumber = payload.phoneNumber;
      candidate.email = payload.email;
      candidate.indeedProfileUrl = payload.indeedUrl;
      candidate.livesIn = payload.livesIn;
      candidate.earmarkedFor = payload.earmarkedFor;
      candidate.keepInMind = payload.keepInMind;
      candidate.tags = payload.tags;
      candidate.notes = payload.notes;
      setDetail(candidate);
    }
    renderFilterOptions();
    renderCandidates();
    setStatus("Candidate details saved.", false, { subtle: true });
  } catch (error) {
    console.error(error);
    setStatus(error?.message || "Could not save candidate details.", true, { autoClear: false });
  } finally {
    detailSaveBusy = false;
    setDetailFormEnabled(Boolean(selectedCandidateId));
  }
}

async function saveCandidateStatus() {
  if (statusUpdateBusy || !statusUpdateSelect) {
    return;
  }
  const selectedStatus = cleanText(statusUpdateSelect.value);
  if (!selectedCandidateId || !selectedStatus) {
    setStatus("Select a candidate and status first.", true);
    return;
  }

  statusUpdateBusy = true;
  if (saveStatusBtn) {
    saveStatusBtn.disabled = true;
  }
  if (statusUpdateSelect) {
    statusUpdateSelect.disabled = true;
  }
  try {
    await updateCandidateStatusById(selectedCandidateId, selectedStatus);
  } catch (error) {
    console.error(error);
    setStatus(error?.message || "Could not update status.", true);
  } finally {
    statusUpdateBusy = false;
    if (saveStatusBtn) {
      saveStatusBtn.disabled = false;
    }
    if (statusUpdateSelect) {
      statusUpdateSelect.disabled = false;
    }
  }
}

function redirectToUnauthorized(pageKey) {
  const page = encodeURIComponent(String(pageKey || "recruitment").trim().toLowerCase());
  window.location.href = `./unauthorized.html?page=${page}`;
}

async function loadRecruitmentCandidates() {
  const payload = await directoryApi.listRecruitment();
  allCandidates = Array.isArray(payload?.items) ? payload.items : [];
  recruitmentStatusOptions = normalizeStatusOptions(payload?.choiceOptions?.status);
  recruitmentOwnerOptions = normalizeOwnerOptions(payload?.choiceOptions?.currentOwner);
  currentUserOwnerChoice = resolveMineOwnerChoice(currentUserEmail, recruitmentOwnerOptions);
  if (mineOnlyFilterInput) {
    mineOnlyFilterInput.disabled = !currentUserOwnerChoice;
    if (!currentUserOwnerChoice) {
      mineOnlyFilterInput.checked = false;
    }
  }
  renderAddRecruitmentStatusOptions();
  if (sharePointListLink) {
    sharePointListLink.href = cleanText(payload?.listUrl) || "#";
  }
  syncAddRecruitmentButton();
  renderFilterOptions();
  renderCandidates();
}

async function createRecruitmentCandidate() {
  const candidateName = cleanText(addCandidateNameInput?.value);
  if (!candidateName) {
    setAddRecruitmentError("Candidate name is required.");
    addCandidateNameInput?.focus();
    return;
  }

  setCreateCandidateBusy(true);
  setAddRecruitmentError("");
  try {
    const result = await directoryApi.createRecruitmentCandidate({
      candidateName,
      status: cleanText(addCandidateStatusSelect?.value) || "Organise Initial Call",
      email: cleanText(addCandidateEmailInput?.value),
      phoneNumber: cleanText(addCandidatePhoneInput?.value),
      livesIn: cleanText(addCandidateLivesInInput?.value),
      location: cleanText(addCandidateJobLocationInput?.value),
      source: cleanText(addCandidateSourceInput?.value),
      indeedUrl: cleanText(addCandidateIndeedUrlInput?.value),
      active: addCandidateActiveInput?.checked !== false,
      updateExistingByPhone: addCandidateUpdateExistingInput?.checked === true,
      notes: cleanText(addCandidateNotesInput?.value),
    });
    await loadRecruitmentCandidates();
    const createdId = cleanText(result?.item?.id);
    const createdCandidate = allCandidates.find((candidate) => cleanText(candidate.id) === createdId) || result?.item || null;
    if (createdCandidate?.id) {
      selectedCandidateId = cleanText(createdCandidate.id);
      setDetail(createdCandidate);
      openCandidateDetail(createdCandidate);
    }
    closeAddRecruitmentModal({ force: true });
    setStatus(
      result?.updatedExisting
        ? `Updated existing candidate by phone match: ${candidateName}.`
        : `Candidate added: ${candidateName}.`
    );
  } catch (error) {
    console.error(error);
    setAddRecruitmentError(error?.message || "Could not add candidate.");
  } finally {
    setCreateCandidateBusy(false);
  }
}

async function init() {
  try {
    const account = await authController.restoreSession();
    if (!account) {
      window.location.href = "./index.html";
      return;
    }

    const profile = await directoryApi.getCurrentUser();
    currentUserEmail = cleanText(profile?.email).toLowerCase();
    const role = String(profile?.role || "").trim().toLowerCase();
    if (!canAccessPage(role, "recruitment")) {
      redirectToUnauthorized("recruitment");
      return;
    }

    renderTopNavigation({ role });
    setStatus("Loading recruitment candidates...");
    await loadRecruitmentCandidates();
    setStatus(`Loaded ${allCandidates.length} candidate(s).`);
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
ownerFilterSelect?.addEventListener("change", renderCandidates);
mineOnlyFilterInput?.addEventListener("change", renderCandidates);
for (const button of stageModeFilterButtons) {
  button.addEventListener("click", () => {
    setStageModeFilter(button.dataset.stageModeFilter);
    renderCandidates();
  });
}
sourceFilterSelect?.addEventListener("change", renderCandidates);
activeFilterSelect?.addEventListener("change", renderCandidates);
sortFilterSelect?.addEventListener("change", renderCandidates);
for (const button of sortHeaderButtons) {
  button.addEventListener("click", () => {
    toggleHeaderSort(button.dataset.sortHeader);
  });
}
toggleRecruitmentToolbarBtn?.addEventListener("click", () => {
  setRecruitmentToolbarVisible(recruitmentToolbarContent?.hidden);
});
addRecruitmentItemBtn?.addEventListener("click", () => {
  openAddRecruitmentModal();
});
addRecruitmentJsonDropZone?.addEventListener("click", () => {
  if (createCandidateBusy) {
    return;
  }
  addRecruitmentJsonFileInput?.click();
});
addRecruitmentJsonDropZone?.addEventListener("keydown", (event) => {
  if (createCandidateBusy) {
    return;
  }
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    addRecruitmentJsonFileInput?.click();
  }
});
addRecruitmentJsonDropZone?.addEventListener("dragover", (event) => {
  event.preventDefault();
  if (!createCandidateBusy) {
    addRecruitmentJsonDropZone.classList.add("is-dragover");
  }
});
addRecruitmentJsonDropZone?.addEventListener("dragleave", () => {
  addRecruitmentJsonDropZone.classList.remove("is-dragover");
});
addRecruitmentJsonDropZone?.addEventListener("drop", async (event) => {
  event.preventDefault();
  addRecruitmentJsonDropZone.classList.remove("is-dragover");
  if (createCandidateBusy) {
    return;
  }
  const file = event.dataTransfer?.files?.[0] || null;
  if (!file) {
    return;
  }
  await handleAddRecruitmentJsonFile(file);
});
addRecruitmentJsonFileInput?.addEventListener("change", async () => {
  const file = addRecruitmentJsonFileInput.files?.[0] || null;
  if (!file || createCandidateBusy) {
    return;
  }
  await handleAddRecruitmentJsonFile(file);
});
addRecruitmentForm?.addEventListener("submit", async (event) => {
  event.preventDefault();
  await createRecruitmentCandidate();
});
addRecruitmentCloseBtn?.addEventListener("click", closeAddRecruitmentModal);
cancelAddRecruitmentBtn?.addEventListener("click", closeAddRecruitmentModal);
candidateDetailCloseBtn?.addEventListener("click", closeCandidateDetail);

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
saveStatusBtn?.addEventListener("click", async () => {
  await saveCandidateStatus();
});
saveDetailBtn?.addEventListener("click", async () => {
  await saveCandidateDetails();
});
candidateDetailEditForm?.addEventListener("submit", async (event) => {
  event.preventDefault();
  await saveCandidateDetails();
});
recruitmentTableBody?.addEventListener("click", async (event) => {
  const saveButton = event.target instanceof Element ? event.target.closest("[data-stage-save]") : null;
  if (saveButton) {
    event.stopPropagation();
    const card = saveButton.closest(".recruitment-stage-card");
    if (!card) {
      return;
    }
    const itemId = cleanText(card.getAttribute("data-item-id"));
    const stageKey = cleanText(card.getAttribute("data-stage-key"));
    const outcome = cleanText(card.querySelector("[data-stage-outcome]")?.value);
    const nextSteps = cleanText(card.querySelector("[data-stage-next-steps]")?.value);
    await saveRecruitmentStage(itemId, stageKey, outcome, nextSteps);
    return;
  }

  const stageCard = event.target instanceof Element ? event.target.closest(".recruitment-stage-card") : null;
  if (stageCard) {
    event.stopPropagation();
  }
}, true);
recruitmentTableBody?.addEventListener("toggle", (event) => {
  const detailsEl = event.target instanceof Element ? event.target.closest(".recruitment-stage-card") : null;
  if (!(detailsEl instanceof HTMLDetailsElement)) {
    return;
  }
  const itemId = cleanText(detailsEl.getAttribute("data-item-id"));
  const stageKey = cleanText(detailsEl.getAttribute("data-stage-key"));
  if (!itemId || !stageKey) {
    return;
  }
  const stateKey = `${itemId}:${stageKey}`;
  if (detailsEl.open) {
    openStageKeys.add(stateKey);
  } else {
    openStageKeys.delete(stateKey);
  }
}, true);
detailInputs.tags?.addEventListener("input", () => {
  renderTagPreview(detailTagsPreview, detailInputs.tags.value);
});
detailInputs.tags?.addEventListener("blur", () => {
  detailInputs.tags.value = normalizeTagString(detailInputs.tags.value);
  renderTagPreview(detailTagsPreview, detailInputs.tags.value);
});
detailInputs.indeedUrl?.addEventListener("input", () => {
  setIndeedButton(detailInputs.indeedUrl.value);
});
detailInputs.indeedUrl?.addEventListener("blur", () => {
  detailInputs.indeedUrl.value = cleanText(detailInputs.indeedUrl.value);
  setIndeedButton(detailInputs.indeedUrl.value);
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
candidateDetailModal?.addEventListener("click", (event) => {
  if (event.target !== candidateDetailModal) {
    return;
  }
  closeCandidateDetail();
});
addRecruitmentModal?.addEventListener("click", (event) => {
  if (event.target !== addRecruitmentModal || createCandidateBusy) {
    return;
  }
  closeAddRecruitmentModal();
});

document.addEventListener("keydown", (event) => {
  if (event.key !== "Escape" || addToOneTouchBusy) {
    return;
  }
  closeStatusQuickMenu();
  closeCandidateDetail();
  if (!oneTouchPickerModal?.hidden) {
    closeOneTouchPicker();
  }
  if (!addRecruitmentModal?.hidden && !createCandidateBusy) {
    closeAddRecruitmentModal();
  }
});

document.addEventListener("click", (event) => {
  if (statusQuickMenu?.hidden) {
    return;
  }
  const target = event.target;
  if (!(target instanceof Node)) {
    return;
  }
  if (statusQuickMenu.contains(target)) {
    return;
  }
  closeStatusQuickMenu();
});

document.addEventListener(
  "scroll",
  (event) => {
    if (statusQuickMenu?.hidden) {
      return;
    }
    const target = event.target;
    if (target instanceof Node && statusQuickMenu?.contains(target)) {
      return;
    }
    closeStatusQuickMenu();
  },
  true
);

signOutBtn?.addEventListener("click", async () => {
  try {
    signOutBtn.disabled = true;
    await authController.signOut();
  } finally {
    window.location.href = "./index.html";
  }
});

setStageModeFilter("all");
syncAddRecruitmentButton();
void init();
